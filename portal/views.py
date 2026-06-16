from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import AuthenticationForm
from django.contrib import messages
from django.db.models import F, Sum
from django.http import HttpResponse
from django.utils import timezone
from functools import wraps
import datetime
import json
import io

# Excel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# PDF
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from .models import Category, Product, Bill, BillItem, StockHistory, Expense
from .forms import StockForm

# 1. Custom Role-Based Decorators
def admin_required(view_func):
    """Decorator for views that checks that the logged-in user is staff or superuser."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('login')
        if request.user.is_staff or request.user.is_superuser:
            return view_func(request, *args, **kwargs)
        # Standard user is redirected to their day billing home
        messages.error(request, "Access Denied: You do not have administrator permissions.")
        return redirect('billing')
    return _wrapped_view


def standard_user_required(view_func):
    """Decorator for views that checks that the logged-in user is NOT staff or superuser (Billing/Counter staff)."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('login')
        if not (request.user.is_staff or request.user.is_superuser):
            return view_func(request, *args, **kwargs)
        # Admins are redirected to their dashboard home
        messages.error(request, "Access Denied: Billing and Stocks screens are for standard billing operators only.")
        return redirect('dashboard')
    return _wrapped_view


# 2. Authentication Views
def login_view(request):
    if request.user.is_authenticated:
        if request.user.is_staff or request.user.is_superuser:
            return redirect('dashboard')
        return redirect('billing')
        
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user is not None:
                login(request, user)
                messages.success(request, f"Welcome back, {username}!")
                if user.is_staff or user.is_superuser:
                    return redirect('dashboard')
                return redirect('billing')
        else:
            messages.error(request, "Invalid username or password.")
    else:
        form = AuthenticationForm()
    return render(request, 'login.html', {'form': form})


@login_required
def logout_view(request):
    logout(request)
    messages.info(request, "You have been logged out successfully.")
    return redirect('login')


# 3. Read-Only Stocks View (For Standard User Only)
@standard_user_required
def stocks_view(request):
    products = Product.objects.all()
    categories = Category.objects.all()
    active_category = request.GET.get('category')
    
    if active_category:
        products = products.filter(category__slug=active_category)
        
    context = {
        'products': products,
        'categories': categories,
        'active_category': active_category,
    }
    return render(request, 'stocks.html', context)


# 4. Day Billing View (For Standard User Only)
@standard_user_required
def billing_view(request):
    products = Product.objects.all()

    if request.method == 'POST':
        cust_name    = request.POST.get('customer_name', '').strip()
        cust_phone   = request.POST.get('customer_phone', '').strip()
        cust_address = request.POST.get('customer_address', '').strip()
        amount_given = float(request.POST.get('amount_given', 0.00))

        # Validate phone number if provided
        if cust_phone:
            if not cust_phone.isdigit() or len(cust_phone) != 10:
                messages.error(request, "Phone number must be exactly 10 digits (numbers only).")
                return render(request, 'billing.html', {'products': products})

        # Collect all item rows sent from the form
        product_ids = request.POST.getlist('product_id[]')
        quantities  = request.POST.getlist('quantity[]')
        rates       = request.POST.getlist('rate[]')

        if not product_ids:
            messages.error(request, "Please add at least one product to the bill.")
            return render(request, 'billing.html', {'products': products})

        # Validate all items first before saving anything
        items_data = []
        grand_total = 0.0
        errors = []

        for i, pid in enumerate(product_ids):
            try:
                qty  = int(quantities[i])
                rate = float(rates[i])
            except (ValueError, IndexError):
                errors.append(f"Row {i+1}: Invalid quantity or rate.")
                continue

            if qty <= 0 or rate < 0:
                errors.append(f"Row {i+1}: Quantity must be > 0 and rate must be >= 0.")
                continue

            try:
                product = Product.objects.get(id=pid)
            except Product.DoesNotExist:
                errors.append(f"Row {i+1}: Product not found.")
                continue

            if product.stock_quantity < qty:
                errors.append(
                    f"'{product.name}': Requested {qty} but only {product.stock_quantity} in stock."
                )
                continue

            item_total = qty * rate
            grand_total += item_total
            items_data.append({'product': product, 'qty': qty, 'rate': rate, 'total': item_total})

        if errors:
            for err in errors:
                messages.error(request, err)
            return render(request, 'billing.html', {'products': products})

        # All valid — create bill header
        amount_to_be_given = amount_given - grand_total
        bill = Bill.objects.create(
            customer_name=cust_name,
            customer_phone=cust_phone or None,
            customer_address=cust_address,
            grand_total=grand_total,
            amount_given=amount_given,
            amount_to_be_given=amount_to_be_given,
        )

        # Create line items and deduct stock
        for item in items_data:
            BillItem.objects.create(
                bill=bill,
                product=item['product'],
                quantity=item['qty'],
                rate=item['rate'],
                total=item['total'],
            )
            item['product'].stock_quantity -= item['qty']
            item['product'].save()

        change_text = (
            f"Change: ₹{amount_to_be_given:.2f}"
            if amount_to_be_given >= 0
            else f"Due: ₹{abs(amount_to_be_given):.2f}"
        )
        messages.success(
            request,
            f"Bill #{bill.id} for {cust_name} created with {len(items_data)} item(s). {change_text}"
        )
        return redirect('billing_records')

    return render(request, 'billing.html', {'products': products})


# 4b. Day Billing Records (For Standard User Only)
@standard_user_required
def billing_records_view(request):
    date_str = request.GET.get('date')
    query_date = timezone.localdate()
    
    if date_str:
        try:
            query_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            pass
            
    # Define start and end of the day in local timezone
    start_of_day = timezone.make_aware(datetime.datetime.combine(query_date, datetime.time.min))
    end_of_day = timezone.make_aware(datetime.datetime.combine(query_date, datetime.time.max))
    
    daily_bills = Bill.objects.filter(created_at__range=(start_of_day, end_of_day)).order_by('-created_at')
    
    context = {
        'daily_bills': daily_bills,
        'today': query_date,
    }
    return render(request, 'billing_records.html', context)



# Helper to fetch user dashboard statistics and daily breakdown
def _get_user_dashboard_data(request):
    today = timezone.localdate()
    start_date_default = today.replace(day=1)
    end_date_default = today

    start_date_str = request.GET.get('start_date')
    end_date_str = request.GET.get('end_date')

    try:
        start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d').date() if start_date_str else start_date_default
    except ValueError:
        start_date = start_date_default

    try:
        end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date() if end_date_str else end_date_default
    except ValueError:
        end_date = end_date_default

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    # Limit to 90 days to prevent rendering lag
    if (end_date - start_date).days > 90:
        start_date = end_date - datetime.timedelta(days=90)

    daily_data = []
    current_day = start_date
    total_sales = 0.0
    total_due = 0.0
    total_expenditure = 0.0

    while current_day <= end_date:
        day_start = timezone.make_aware(datetime.datetime.combine(current_day, datetime.time.min))
        day_end = timezone.make_aware(datetime.datetime.combine(current_day, datetime.time.max))
        
        # Sales
        day_bills = Bill.objects.filter(created_at__range=(day_start, day_end))
        day_sales = float(day_bills.aggregate(total=Sum('grand_total'))['total'] or 0.00)
        
        # Due
        day_due = abs(float(day_bills.filter(amount_to_be_given__lt=0).aggregate(total=Sum('amount_to_be_given'))['total'] or 0.00))
        
        # Expenditure
        day_hist = StockHistory.objects.filter(date_entered=current_day, qty_added__gt=0)
        day_hist_total = sum(float(h.qty_added * h.product.purchase_price) for h in day_hist.select_related('product'))
        
        day_init = Product.objects.filter(stock_entered=current_day)
        day_init_total = sum(float(p.initial_quantity * p.purchase_price) for p in day_init)
        
        day_manual_exp = float(Expense.objects.filter(date_paid=current_day).aggregate(total=Sum('amount'))['total'] or 0.00)
        
        day_exp = day_hist_total + day_init_total + day_manual_exp
        
        total_sales += day_sales
        total_due += day_due
        total_expenditure += day_exp
        
        daily_data.append({
            'date': current_day,
            'display_date': current_day.strftime('%b %d, %Y'),
            'sales': day_sales,
            'due': day_due,
            'expenditure': day_exp,
            'net': day_sales - day_exp
        })
        
        current_day += datetime.timedelta(days=1)

    return daily_data, start_date, end_date, total_sales, total_due, total_expenditure



# 4bc. User-Side Dashboard Excel Export (For All Authenticated Users)
@login_required
def user_dashboard_excel(request):
    daily_data, start_date, end_date, total_sales, total_due, total_expenditure = _get_user_dashboard_data(request)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Accounts Report"

    # Styles
    header_fill = PatternFill("solid", fgColor="2563EB") # Royal Blue
    header_font = Font(bold=True, color="FFFFFF", size=11)
    summary_font = Font(bold=True, size=11)
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws.merge_cells("A1:E1")
    ws["A1"] = "Rukmini Enterprises — Accounts Report"
    ws["A1"].font = Font(bold=True, size=14, color="1E293B")
    ws["A1"].alignment = center

    ws.merge_cells("A2:E2")
    ws["A2"] = f"Period: {start_date.strftime('%d %b %Y')} to {end_date.strftime('%d %b %Y')}"
    ws["A2"].font = Font(size=10, color="64748B")
    ws["A2"].alignment = center

    ws.append([]) # empty row

    # Summary Metrics Table
    ws.append(["Summary Metrics", "", "", "", ""])
    ws.merge_cells("A4:E4")
    ws["A4"].font = Font(bold=True, size=12)
    ws["A4"].fill = PatternFill("solid", fgColor="F3F4F6")

    ws.append(["Total Sales Done", "Total Outstanding Dues", "Total Expenditure", "Net Balance", ""])
    ws.merge_cells("D5:E5")
    for col in range(1, 6):
        ws.cell(row=5, column=col).font = bold
        ws.cell(row=5, column=col).border = border

    ws.append([total_sales, total_due, total_expenditure, total_sales - total_expenditure, ""])
    ws.merge_cells("D6:E6")
    for col in range(1, 6):
        cell = ws.cell(row=6, column=col)
        cell.font = summary_font
        cell.border = border
        if col == 4:
            if (total_sales - total_expenditure) >= 0:
                cell.font = Font(bold=True, color="15803D") # green
            else:
                cell.font = Font(bold=True, color="B91C1C") # red

    ws.append([]) # empty row

    # Headers
    headers = ["Date", "Sales Done (INR)", "Outstanding Dues (INR)", "Expenditure (INR)", "Net Balance (INR)"]
    ws.append(headers)
    for col, cell in enumerate(ws[8], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for d in daily_data:
        ws.append([
            d['display_date'],
            d['sales'],
            d['due'],
            d['expenditure'],
            d['net']
        ])
        row = ws.max_row
        for col_idx, cell in enumerate(ws[row], 1):
            cell.border = border
            if col_idx > 1:
                cell.alignment = right
                cell.number_format = '₹#,##0.00'
            else:
                cell.alignment = center

    # Column widths
    col_widths = [16, 20, 24, 20, 20]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    filename = f"accounts_report_{start_date}_to_{end_date}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response


# 4bd. User-Side Dashboard PDF Export (For All Authenticated Users)
@login_required
def user_dashboard_pdf(request):
    daily_data, start_date, end_date, total_sales, total_due, total_expenditure = _get_user_dashboard_data(request)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    elements = []

    title_style = ParagraphStyle('title', fontSize=15, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#1E293B'), alignment=TA_CENTER, spaceAfter=3)
    sub_style   = ParagraphStyle('sub',   fontSize=9,  fontName='Helvetica',
                                  textColor=colors.HexColor('#64748B'), alignment=TA_CENTER, spaceAfter=12)
    label_style = ParagraphStyle('lbl',   fontSize=9,  fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#475569'), alignment=TA_CENTER)
    val_style   = ParagraphStyle('val',   fontSize=12, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#0F172A'), alignment=TA_CENTER)

    elements.append(Paragraph("Rukmini Enterprises", title_style))
    elements.append(Paragraph(f"Monthly Accounts Report (Period: {start_date.strftime('%d %b %Y')} to {end_date.strftime('%d %b %Y')})", sub_style))

    # Summary table
    summary_data = [
        [
            Paragraph("Total Sales Done", label_style),
            Paragraph("Total Outstanding Dues", label_style),
            Paragraph("Total Expenditure", label_style),
            Paragraph("Net Balance", label_style)
        ],
        [
            Paragraph(f"₹{total_sales:,.2f}", val_style),
            Paragraph(f"₹{total_due:,.2f}", val_style),
            Paragraph(f"₹{total_expenditure:,.2f}", val_style),
            Paragraph(f"₹{total_sales - total_expenditure:,.2f}", ParagraphStyle('net_val', parent=val_style, textColor=colors.HexColor('#15803D') if (total_sales - total_expenditure) >= 0 else colors.HexColor('#B91C1C')))
        ]
    ]

    sum_table = Table(summary_data, colWidths=[4.5*cm, 4.5*cm, 4.5*cm, 4.5*cm])
    sum_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#E2E8F0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    elements.append(sum_table)
    elements.append(Spacer(1, 15))

    # Daily Ledger details
    th_style = ParagraphStyle('th', fontSize=8, fontName='Helvetica-Bold', textColor=colors.white, alignment=TA_CENTER)
    td_style = ParagraphStyle('td', fontSize=8, fontName='Helvetica', textColor=colors.HexColor('#334155'))
    td_num_style = ParagraphStyle('td_num', fontSize=8, fontName='Helvetica', textColor=colors.HexColor('#334155'), alignment=TA_RIGHT)

    table_data = [[
        Paragraph("Date", th_style),
        Paragraph("Sales Done", th_style),
        Paragraph("Outstanding Dues", th_style),
        Paragraph("Expenditure", th_style),
        Paragraph("Net Balance", th_style),
    ]]

    for d in daily_data:
        table_data.append([
            Paragraph(d['display_date'], td_style),
            Paragraph(f"₹{d['sales']:,.2f}", td_num_style),
            Paragraph(f"₹{d['due']:,.2f}", td_num_style),
            Paragraph(f"₹{d['expenditure']:,.2f}", td_num_style),
            Paragraph(f"₹{d['net']:,.2f}", ParagraphStyle('net_td', parent=td_num_style, fontName='Helvetica-Bold', textColor=colors.HexColor('#16A34A') if d['net'] >= 0 else colors.HexColor('#DC2626'))),
        ])

    ledger_table = Table(table_data, colWidths=[3.5*cm, 3.6*cm, 3.7*cm, 3.6*cm, 3.6*cm])
    ledger_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('TOPPADDING', (0,0), (-1,0), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0,1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,1), (-1,-1), 5),
    ]))
    elements.append(ledger_table)

    doc.build(elements)
    pdf_val = buffer.getvalue()
    buffer.close()

    response = HttpResponse(content_type='application/pdf')
    filename = f"accounts_report_{start_date}_to_{end_date}.pdf"
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    response.write(pdf_val)
    return response


# 4c. Clear Due View (For Standard User Only)
@standard_user_required
def clear_due_view(request, pk):
    bill = get_object_or_404(Bill, pk=pk)
    if request.method == 'POST':
        clear_amount_str = request.POST.get('clear_amount')
        if clear_amount_str:
            try:
                clear_amount = float(clear_amount_str)
                if clear_amount > 0:
                    max_due = abs(float(bill.amount_to_be_given))
                    # Prevent floating point inaccuracies and overpaying
                    if clear_amount > round(max_due, 2):
                        clear_amount = max_due
                    
                    bill.amount_given = float(bill.amount_given) + clear_amount
                    bill.amount_to_be_given = float(bill.amount_given) - float(bill.total)
                    bill.save()
                    
                    if bill.amount_to_be_given >= 0:
                        messages.success(request, f"Payment of ₹{clear_amount:.2f} received. Bill #{bill.id} is now fully cleared!")
                    else:
                        messages.success(request, f"Payment of ₹{clear_amount:.2f} received for Bill #{bill.id}. Remaining due: ₹{abs(bill.amount_to_be_given):.2f}")
                else:
                    messages.error(request, "Payment amount must be greater than zero.")
            except ValueError:
                messages.error(request, "Invalid payment amount format.")
        else:
            messages.error(request, "No payment amount was specified.")
            
        local_created_at = timezone.localtime(bill.created_at)
        bill_date_str = local_created_at.strftime('%Y-%m-%d')
        return redirect(f"/billing/records/?date={bill_date_str}")
    return redirect('billing_records')


# 5. Dashboard View (Admin Only)
@admin_required
def dashboard_view(request):
    daily_data, start_date, end_date, total_sales, total_due, total_expenditure = _get_user_dashboard_data(request)

    chart_labels = [d['display_date'] for d in daily_data]
    chart_sales = [d['sales'] for d in daily_data]
    chart_due = [d['due'] for d in daily_data]
    chart_exp = [d['expenditure'] for d in daily_data]

    context = {
        'start_date': start_date,
        'end_date': end_date,
        'total_sales': total_sales,
        'total_due': total_due,
        'total_expenditure': total_expenditure,
        'daily_data': list(reversed(daily_data)), # Show most recent first in table
        'chart_labels': json.dumps(chart_labels),
        'chart_sales': json.dumps(chart_sales),
        'chart_due': json.dumps(chart_due),
        'chart_exp': json.dumps(chart_exp),
        'is_admin': True,
    }
    return render(request, 'user_dashboard.html', context)


# 5b. Admin Expenses List/Manage View (Admin Only)
@admin_required
def admin_expenses_view(request):
    # Handle POST requests (Create or Delete)
    if request.method == 'POST':
        # 1. Handle Deletion
        delete_id = request.POST.get('delete_id')
        if delete_id:
            expense_to_delete = get_object_or_404(Expense, pk=delete_id)
            comp = expense_to_delete.company_name
            amt = expense_to_delete.amount
            expense_to_delete.delete()
            messages.warning(request, f"Expense of ₹{amt:.2f} to '{comp}' was deleted successfully.")
            return redirect('admin_expenses')
            
        # 2. Handle Creation
        company_name = request.POST.get('company_name', '').strip()
        amount_str   = request.POST.get('amount')
        date_paid_str = request.POST.get('date_paid')
        description  = request.POST.get('description', '').strip()
        
        if company_name and amount_str:
            try:
                amount = float(amount_str)
                if amount <= 0:
                    messages.error(request, "Amount must be greater than zero.")
                else:
                    date_paid = timezone.localdate()
                    if date_paid_str:
                        date_paid = datetime.datetime.strptime(date_paid_str, '%Y-%m-%d').date()
                        
                    Expense.objects.create(
                        company_name=company_name,
                        amount=amount,
                        description=description,
                        date_paid=date_paid,
                        recorded_by=request.user
                    )
                    messages.success(request, f"Expense of ₹{amount:.2f} paid to '{company_name}' recorded successfully.")
                    return redirect('admin_expenses')
            except ValueError:
                messages.error(request, "Invalid amount or date format.")
        else:
            messages.error(request, "Please fill in all required fields.")
            
    # Search and date filters
    company_query = request.GET.get('company', '').strip()
    start_date_str = request.GET.get('start_date')
    end_date_str = request.GET.get('end_date')
    
    today = timezone.localdate()
    start_date = today.replace(day=1)
    end_date = today
    
    try:
        if start_date_str:
            start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d').date()
    except ValueError:
        pass
        
    try:
        if end_date_str:
            end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date()
    except ValueError:
        pass

    expenses = Expense.objects.filter(date_paid__range=(start_date, end_date))
    
    if company_query:
        expenses = expenses.filter(company_name__icontains=company_query)
        
    expenses = expenses.order_by('-date_paid', '-created_at')
    
    # Calculate sum
    total_expense = float(expenses.aggregate(total=Sum('amount'))['total'] or 0.00)

    context = {
        'expenses': expenses,
        'company_query': company_query,
        'start_date': start_date,
        'end_date': end_date,
        'total_expense': total_expense,
        'today': today,
    }
    return render(request, 'admin_expenses.html', context)


# 6. Stock & Stock 1 Views (Admin Only)
@admin_required
def stock_add_view(request):
    categories = Category.objects.all()
    products = Product.objects.all().order_by('-id')
    
    if request.method == 'POST':
        form = StockForm(request.POST, request.FILES)
        if form.is_valid():
            product = form.save(commit=False)
            product.stock_quantity = product.initial_quantity
            product.save()
            messages.success(request, f"Stock item '{product.name}' was added successfully.")
            return redirect('stock_add')
    else:
        form = StockForm()
        
    context = {
        'form': form,
        'products': products,
        'categories': categories,
    }
    return render(request, 'stock.html', context)


@admin_required
def stock1_view(request):
    products = Product.objects.all().order_by('-id')
    return render(request, 'stock1.html', {'products': products})


@admin_required
def stock_edit_view(request, pk):
    product = get_object_or_404(Product, pk=pk)
    categories = Category.objects.all()
    products = Product.objects.all().order_by('-id')
    
    if request.method == 'POST':
        old_initial_qty = product.initial_quantity
        form = StockForm(request.POST, request.FILES, instance=product)
        if form.is_valid():
            product = form.save(commit=False)
            qty_diff = product.initial_quantity - old_initial_qty
            if qty_diff != 0:
                product.stock_quantity = max(0, product.stock_quantity + qty_diff)
            product.save()

            # ── Log stock history whenever qty is changed ──
            if qty_diff != 0:
                StockHistory.objects.create(
                    product=product,
                    date_entered=product.stock_entered or timezone.localdate(),
                    qty_added=qty_diff,
                    qty_after=product.stock_quantity,
                )

            messages.success(request, f"Stock item '{product.name}' was updated successfully.")
            return redirect('stock_add')
    else:
        form = StockForm(instance=product)
        # Default stock_entered to today so admin just confirms the current update date
        form.initial['stock_entered'] = timezone.localdate()
        
    stock_history = product.stock_history.all()

    context = {
        'form': form,
        'products': products,
        'categories': categories,
        'product': product,
        'is_edit': True,
        'stock_history': stock_history,
    }
    return render(request, 'stock.html', context)


@admin_required
def stock_delete_view(request, pk):
    product = get_object_or_404(Product, pk=pk)
    if request.method == 'POST':
        name = product.name
        product.delete()
        messages.warning(request, f"Stock item '{name}' was deleted.")
    
    referer = request.META.get('HTTP_REFERER')
    if referer and 'stock1' in referer:
        return redirect('stock1')
    return redirect('stock_add')


@admin_required
def revenue_report_view(request):
    preset = request.GET.get('preset', 'monthly')
    start_date_str = request.GET.get('start_date')
    end_date_str = request.GET.get('end_date')
    
    today = timezone.localdate()
    start_date = today
    end_date = today
    
    if preset == 'daily':
        start_date = today
        end_date = today
    elif preset == 'monthly':
        start_date = today.replace(day=1)
        end_date = today
    elif preset == '3months':
        start_date = today - datetime.timedelta(days=90)
        end_date = today
    elif preset == '6months':
        start_date = today - datetime.timedelta(days=180)
        end_date = today
    elif preset == '9months':
        start_date = today - datetime.timedelta(days=270)
        end_date = today
    elif preset == 'yearly':
        start_date = today - datetime.timedelta(days=365)
        end_date = today
    elif preset == 'custom':
        if start_date_str and end_date_str:
            try:
                start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
                end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
            except ValueError:
                pass
    else:
        # Default fallback
        preset = 'monthly'
        start_date = today.replace(day=1)
        end_date = today

    # Ensure start_date is not after end_date
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    # Convert local dates to timezone-aware datetime limits
    start_dt = timezone.make_aware(datetime.datetime.combine(start_date, datetime.time.min))
    end_dt = timezone.make_aware(datetime.datetime.combine(end_date, datetime.time.max))

    # Query matching bills
    bills = Bill.objects.filter(created_at__range=(start_dt, end_dt)).order_by('-created_at')

    # Calculate metrics
    total_revenue = float(bills.aggregate(total=Sum('grand_total'))['total'] or 0.0)
    total_sales = bills.count()
    total_paid = float(bills.aggregate(total_paid=Sum('amount_given'))['total_paid'] or 0.0)
    avg_order_value = total_revenue / total_sales if total_sales > 0 else 0.0
    total_dues = float(sum(abs(b.amount_to_be_given) for b in bills if b.amount_to_be_given < 0))

    # Group revenue by date for Chart.js trend lines
    chart_labels = []
    chart_values = []
    
    delta = end_date - start_date
    if delta.days <= 45:
        # Group by day
        day_map = {}
        curr = start_date
        while curr <= end_date:
            day_map[curr] = 0.0
            curr += datetime.timedelta(days=1)
            
        for bill in bills:
            local_bill_date = timezone.localtime(bill.created_at).date()
            if local_bill_date in day_map:
                day_map[local_bill_date] += float(bill.grand_total)
                
        # Format labels nicely
        for d in sorted(day_map.keys()):
            chart_labels.append(d.strftime("%b %d"))
            chart_values.append(day_map[d])
    else:
        # Group by month
        month_map = {}
        curr = start_date
        while curr <= end_date:
            m_key = (curr.year, curr.month)
            month_map[m_key] = 0.0
            if curr.month == 12:
                curr = datetime.date(curr.year + 1, 1, 1)
            else:
                curr = datetime.date(curr.year, curr.month + 1, 1)
                
        # Fill data
        for bill in bills:
            local_bill_dt = timezone.localtime(bill.created_at)
            m_key = (local_bill_dt.year, local_bill_dt.month)
            if m_key in month_map:
                month_map[m_key] += float(bill.grand_total)
                
        # Format labels nicely
        for m in sorted(month_map.keys()):
            temp_date = datetime.date(m[0], m[1], 1)
            chart_labels.append(temp_date.strftime("%b %Y"))
            chart_values.append(month_map[m])

    context = {
        'preset': preset,
        'start_date': start_date,
        'end_date': end_date,
        'bills': bills,
        'total_revenue': total_revenue,
        'total_sales': total_sales,
        'total_paid': total_paid,
        'avg_order_value': avg_order_value,
        'total_dues': total_dues,
        'chart_labels': json.dumps(chart_labels),
        'chart_values': json.dumps(chart_values),
    }
    return render(request, 'revenue_report.html', context)


# ─────────────────────────────────────────────────────────────
# DOWNLOAD VIEWS
# ─────────────────────────────────────────────────────────────

def _get_report_bills(request):
    """Shared helper: parse date params and return filtered bills + metadata."""
    preset = request.GET.get('preset', 'monthly')
    start_date_str = request.GET.get('start_date')
    end_date_str = request.GET.get('end_date')
    today = timezone.localdate()
    start_date = today
    end_date = today

    if preset == 'daily':
        start_date = end_date = today
    elif preset == 'monthly':
        start_date = today.replace(day=1)
    elif preset == '3months':
        start_date = today - datetime.timedelta(days=90)
    elif preset == '6months':
        start_date = today - datetime.timedelta(days=180)
    elif preset == '9months':
        start_date = today - datetime.timedelta(days=270)
    elif preset == 'yearly':
        start_date = today - datetime.timedelta(days=365)
    elif preset == 'custom' and start_date_str and end_date_str:
        try:
            start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
            end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
        except ValueError:
            pass
    else:
        preset = 'monthly'
        start_date = today.replace(day=1)

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    start_dt = timezone.make_aware(datetime.datetime.combine(start_date, datetime.time.min))
    end_dt = timezone.make_aware(datetime.datetime.combine(end_date, datetime.time.max))
    bills = Bill.objects.filter(created_at__range=(start_dt, end_dt)).order_by('created_at').prefetch_related('items__product')
    return bills, start_date, end_date


# ── Revenue Report: Excel Download ──────────────────────────
@admin_required
def revenue_report_excel(request):
    bills, start_date, end_date = _get_report_bills(request)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Revenue Report"

    # Styles
    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    sub_fill    = PatternFill("solid", fgColor="EEF2FF")
    bold        = Font(bold=True)
    center      = Alignment(horizontal="center", vertical="center")
    right       = Alignment(horizontal="right")
    thin        = Side(style="thin", color="D1D5DB")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws.merge_cells("A1:H1")
    ws["A1"] = f"Rukmini Enterprises — Revenue Report"
    ws["A1"].font = Font(bold=True, size=14, color="1E293B")
    ws["A1"].alignment = center

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Period: {start_date.strftime('%d %b %Y')} to {end_date.strftime('%d %b %Y')}"
    ws["A2"].font = Font(size=10, color="64748B")
    ws["A2"].alignment = center

    ws.append([])

    # Summary row
    total_revenue = float(bills.aggregate(t=Sum('grand_total'))['t'] or 0)
    total_paid    = float(bills.aggregate(t=Sum('amount_given'))['t'] or 0)
    total_dues    = float(sum(abs(b.amount_to_be_given) for b in bills if b.amount_to_be_given < 0))
    total_sales   = bills.count()

    ws.append(["Total Sales", total_sales, "", "Total Revenue", f"Rs.{total_revenue:.2f}", "", "Amount Collected", f"Rs.{total_paid:.2f}"])
    ws.append(["Amount Balance", f"Rs.{total_dues:.2f}"])
    for cell in ws[4] + ws[5]:
        cell.font = bold
        cell.fill = sub_fill
    ws.append([])

    # Column headers — one row per bill item
    headers = ["Bill#", "Date", "Customer", "Product", "Qty", "Rate (Rs.)", "Item Total (Rs.)", "Bill Total (Rs.)", "Paid (Rs.)", "Balance (Rs.)", "Status"]
    ws.append(headers)
    for col, cell in enumerate(ws[7], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # Data rows — one row per BillItem
    for bill in bills:
        status  = "Paid" if bill.amount_to_be_given >= 0 else "Due"
        balance = float(abs(bill.amount_to_be_given)) if bill.amount_to_be_given < 0 else 0
        items   = list(bill.items.select_related('product').all())
        for idx, item in enumerate(items):
            ws.append([
                bill.id if idx == 0 else "",
                timezone.localtime(bill.created_at).strftime("%d %b %Y %I:%M %p") if idx == 0 else "",
                bill.customer_name if idx == 0 else "",
                item.product.name if item.product else "Deleted",
                item.quantity,
                float(item.rate),
                float(item.total),
                float(bill.grand_total) if idx == 0 else "",
                float(bill.amount_given) if idx == 0 else "",
                balance if idx == 0 else "",
                status if idx == 0 else "",
            ])
            row = ws.max_row
            for cell in ws[row]:
                cell.border = border
                cell.alignment = Alignment(vertical="center")

    # Column widths
    col_widths = [6, 22, 22, 24, 6, 12, 14, 14, 14, 10]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    filename = f"revenue_report_{start_date}_{end_date}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response


# ── Revenue Report: PDF Download ─────────────────────────────
@admin_required
def revenue_report_pdf(request):
    bills, start_date, end_date = _get_report_bills(request)

    buffer = io.BytesIO()
    # A4 landscape usable width ≈ 27.7cm after margins
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    elements = []

    title_style = ParagraphStyle('title', fontSize=15, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#1E293B'), alignment=TA_CENTER, spaceAfter=3)
    sub_style   = ParagraphStyle('sub',   fontSize=9,  fontName='Helvetica',
                                  textColor=colors.HexColor('#64748B'), alignment=TA_CENTER, spaceAfter=12)

    elements.append(Paragraph("Rukmini Enterprises - Revenue Report", title_style))
    elements.append(Paragraph(f"Period: {start_date.strftime('%d %b %Y')} to {end_date.strftime('%d %b %Y')}", sub_style))

    # Summary box — 4 label+value pairs
    total_revenue = float(bills.aggregate(t=Sum('grand_total'))['t'] or 0)
    total_paid    = float(bills.aggregate(t=Sum('amount_given'))['t'] or 0)
    total_dues    = float(sum(abs(b.amount_to_be_given) for b in bills if b.amount_to_be_given < 0))
    total_sales   = bills.count()

    summary_data = [
        ["Total Sales", str(total_sales),
         "Total Revenue", f"Rs. {total_revenue:,.2f}",
         "Amount Collected", f"Rs. {total_paid:,.2f}",
         "Amount Balance", f"Rs. {total_dues:,.2f}"],
    ]
    # 8 equal columns across full width
    sw = [3.46*cm] * 8
    summary_table = Table(summary_data, colWidths=sw)
    summary_table.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), colors.HexColor('#EEF2FF')),
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,-1), 8.5),
        ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('BOX',           (0,0), (-1,-1), 0.6, colors.HexColor('#C7D2FE')),
        ('INNERGRID',     (0,0), (-1,-1), 0.4, colors.HexColor('#C7D2FE')),
        ('TOPPADDING',    (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING',   (0,0), (-1,-1), 6),
        ('RIGHTPADDING',  (0,0), (-1,-1), 6),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.35*cm))

    # Detail table
    col_headers = ["Bill#", "Date", "Customer", "Product", "Qty", "Rate", "Item Total", "Bill Total", "Paid", "Balance", "Status"]
    col_widths_pdf = [1.0*cm, 2.6*cm, 3.6*cm, 3.8*cm, 1.2*cm, 2.6*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 1.6*cm]
    data = [col_headers]
    for bill in bills:
        status  = "Paid" if bill.amount_to_be_given >= 0 else "Due"
        balance = f"Rs.{abs(float(bill.amount_to_be_given)):,.2f}" if bill.amount_to_be_given < 0 else "-"
        items   = list(bill.items.select_related('product').all())
        for idx, item in enumerate(items):
            data.append([
                str(bill.id) if idx == 0 else "",
                timezone.localtime(bill.created_at).strftime("%d %b %y") if idx == 0 else "",
                bill.customer_name[:18] if idx == 0 else "",
                (item.product.name[:20] if item.product else "Deleted"),
                str(item.quantity),
                f"Rs.{float(item.rate):,.2f}",
                f"Rs.{float(item.total):,.2f}",
                f"Rs.{float(bill.grand_total):,.2f}" if idx == 0 else "",
                f"Rs.{float(bill.amount_given):,.2f}" if idx == 0 else "",
                balance if idx == 0 else "",
                status if idx == 0 else "",
            ])

    if len(data) == 1:
        data.append(["No records in this period."] + [""]*10)

    tbl = Table(data, colWidths=col_widths_pdf, repeatRows=1)
    tbl.setStyle(TableStyle([
        # ── Header row ──
        ('BACKGROUND',    (0,0), (-1,0),  colors.HexColor('#4F46E5')),
        ('TEXTCOLOR',     (0,0), (-1,0),  colors.white),
        ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0),  8),
        ('ALIGN',         (0,0), (-1,0),  'CENTER'),
        ('VALIGN',        (0,0), (-1,0),  'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,0),  6),
        ('BOTTOMPADDING', (0,0), (-1,0),  6),
        # ── Data rows ──
        ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1), (-1,-1), 7.5),
        ('VALIGN',        (0,1), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,1), (-1,-1), 5),
        # Left-align text columns: #, Date, Customer, Product
        ('ALIGN',         (0,1), (0,-1),  'CENTER'),   # #
        ('ALIGN',         (1,1), (1,-1),  'LEFT'),     # Date
        ('ALIGN',         (2,1), (2,-1),  'LEFT'),     # Customer
        ('ALIGN',         (3,1), (3,-1),  'LEFT'),     # Product
        # Center Qty
        ('ALIGN',         (4,1), (4,-1),  'CENTER'),   # Qty
        # Right-align numeric columns
        ('ALIGN',         (5,1), (8,-1),  'RIGHT'),    # Rate, Total, Paid, Balance
        # Center Status
        ('ALIGN',         (9,1), (9,-1),  'CENTER'),   # Status
        # Alternating row colors
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        # Borders
        ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor('#E2E8F0')),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
    ]))
    elements.append(tbl)

    doc.build(elements)
    buffer.seek(0)
    filename = f"revenue_report_{start_date}_{end_date}.pdf"
    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


# ── Stock Report: Excel Download ─────────────────────────────
@admin_required
def stock_report_excel(request):
    products = Product.objects.all().select_related('category').order_by('category__name', 'name')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock Report"

    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    low_fill    = PatternFill("solid", fgColor="FEE2E2")
    bold        = Font(bold=True)
    center      = Alignment(horizontal="center", vertical="center")
    thin        = Side(style="thin", color="D1D5DB")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:I1")
    ws["A1"] = "Rukmini Enterprises — Stock Report"
    ws["A1"].font = Font(bold=True, size=14, color="1E293B")
    ws["A1"].alignment = center

    ws.merge_cells("A2:I2")
    ws["A2"] = f"Generated: {timezone.localdate().strftime('%d %b %Y')}"
    ws["A2"].font = Font(size=10, color="64748B")
    ws["A2"].alignment = center
    ws.append([])

    headers = ["#", "Product Name", "Category", "Company", "Purchase Price", "Selling Price", "Stock Qty", "Status"]
    ws.append(headers)
    for col, cell in enumerate(ws[4], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for i, p in enumerate(products, 1):
        status = "Low Stock" if p.is_low_stock else "OK"
        ws.append([
            i,
            p.name,
            p.category.name if p.category else "—",
            p.company_name or "—",
            float(p.purchase_price),
            float(p.selling_price),
            p.stock_quantity,
            status,
        ])
        row = ws.max_row
        for cell in ws[row]:
            cell.border = border
            cell.alignment = Alignment(vertical="center")
        if p.is_low_stock:
            for cell in ws[row]:
                cell.fill = low_fill

    col_widths = [5, 30, 18, 22, 16, 16, 10, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="stock_report_{timezone.localdate()}.xlsx"'
    wb.save(response)
    return response


# ── Stock Report: PDF Download ────────────────────────────────
@admin_required
def stock_report_pdf(request):
    products = Product.objects.all().select_related('category').order_by('category__name', 'name')

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    elements = []

    title_style = ParagraphStyle('title2', fontSize=15, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#1E293B'), alignment=TA_CENTER, spaceAfter=3)
    sub_style   = ParagraphStyle('sub2',   fontSize=9,  fontName='Helvetica',
                                  textColor=colors.HexColor('#64748B'), alignment=TA_CENTER, spaceAfter=12)

    elements.append(Paragraph("Rukmini Enterprises - Stock Report", title_style))
    elements.append(Paragraph(f"Generated: {timezone.localdate().strftime('%d %b %Y')}", sub_style))

    # Columns: #, Product Name, Category, Company, Purchase Price, Selling Price, Stock, Status
    # Total width = 27.7cm
    col_widths_pdf = [1.0*cm, 7.0*cm, 4.2*cm, 4.8*cm, 3.4*cm, 3.4*cm, 2.0*cm, 2.0*cm]

    col_headers = ["#", "Product Name", "Category", "Company", "Purchase Price", "Selling Price", "Stock", "Status"]
    data = [col_headers]
    for i, p in enumerate(products, 1):
        status = "Low Stock" if p.is_low_stock else "OK"
        data.append([
            str(i),
            p.name[:32],
            p.category.name[:16] if p.category else "-",
            (p.company_name or "-")[:22],
            f"Rs.{float(p.purchase_price):,.2f}",
            f"Rs.{float(p.selling_price):,.2f}",
            str(p.stock_quantity),
            status,
        ])

    if len(data) == 1:
        data.append(["No stock items found."] + [""] * 7)

    tbl = Table(data, colWidths=col_widths_pdf, repeatRows=1)
    tbl.setStyle(TableStyle([
        # ── Header row ──
        ('BACKGROUND',    (0,0), (-1,0),  colors.HexColor('#4F46E5')),
        ('TEXTCOLOR',     (0,0), (-1,0),  colors.white),
        ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0),  8),
        ('ALIGN',         (0,0), (-1,0),  'CENTER'),
        ('VALIGN',        (0,0), (-1,0),  'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,0),  6),
        ('BOTTOMPADDING', (0,0), (-1,0),  6),
        # ── Data rows ──
        ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1), (-1,-1), 7.5),
        ('VALIGN',        (0,1), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,1), (-1,-1), 5),
        # Center # column
        ('ALIGN',         (0,1), (0,-1),  'CENTER'),
        # Left-align text columns: Name, Category, Company
        ('ALIGN',         (1,1), (3,-1),  'LEFT'),
        # Right-align price columns
        ('ALIGN',         (4,1), (5,-1),  'RIGHT'),    # Purchase Price, Selling Price
        # Center Stock and Status
        ('ALIGN',         (6,1), (7,-1),  'CENTER'),
        # Alternating row colors
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        # Borders
        ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor('#E2E8F0')),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
    ]))
    elements.append(tbl)

    doc.build(elements)
    buffer.seek(0)
    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="stock_report_{timezone.localdate()}.pdf"'
    return response


# ── Stock History: Excel Download ────────────────────────────
@admin_required
def stock_history_excel(request, pk):
    product = get_object_or_404(Product, pk=pk)
    history = product.stock_history.all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock History"

    # Styles
    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    add_fill    = PatternFill("solid", fgColor="D1FAE5")   # green tint for additions
    rem_fill    = PatternFill("solid", fgColor="FEE2E2")   # red tint for removals
    bold        = Font(bold=True)
    center      = Alignment(horizontal="center", vertical="center")
    thin        = Side(style="thin", color="D1D5DB")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title block
    ws.merge_cells("A1:E1")
    ws["A1"] = f"Rukmini Enterprises — Stock Update History"
    ws["A1"].font = Font(bold=True, size=14, color="1E293B")
    ws["A1"].alignment = center

    ws.merge_cells("A2:E2")
    ws["A2"] = f"Product: {product.name}  |  Company: {product.company_name or '—'}  |  Category: {product.category.name if product.category else '—'}"
    ws["A2"].font = Font(size=10, color="64748B")
    ws["A2"].alignment = center

    ws.merge_cells("A3:E3")
    ws["A3"] = f"Generated: {timezone.localdate().strftime('%d %b %Y')}  |  Current Stock: {product.stock_quantity} units"
    ws["A3"].font = Font(size=10, color="64748B")
    ws["A3"].alignment = center
    ws.append([])

    # Column headers
    headers = ["#", "Date of Stock Entry", "Qty Added / Removed", "Total Stock After Update", "Logged At"]
    ws.append(headers)
    for col, cell in enumerate(ws[5], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # Data rows
    for i, entry in enumerate(history, 1):
        ws.append([
            i,
            entry.date_entered.strftime("%d %b %Y"),
            f"+{entry.qty_added}" if entry.qty_added > 0 else str(entry.qty_added),
            entry.qty_after,
            timezone.localtime(entry.recorded_at).strftime("%d %b %Y, %I:%M %p"),
        ])
        row = ws.max_row
        fill = add_fill if entry.qty_added > 0 else rem_fill
        for cell in ws[row]:
            cell.border = border
            cell.alignment = Alignment(vertical="center", horizontal="center")
        # Colour the qty column
        ws.cell(row=row, column=3).fill = fill

    if not history.exists():
        ws.append(["", "No history records found.", "", "", ""])

    # Column widths
    for i, w in enumerate([5, 22, 22, 26, 28], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    safe_name = product.name.replace(" ", "_")[:30]
    response["Content-Disposition"] = f'attachment; filename="stock_history_{safe_name}_{timezone.localdate()}.xlsx"'
    wb.save(response)
    return response


# ── Stock History: PDF Download ────────────────────────────
@admin_required
def stock_history_pdf(request, pk):
    product = get_object_or_404(Product, pk=pk)
    history = product.stock_history.all()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    elements = []

    title_style = ParagraphStyle('htitle', fontSize=15, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#1E293B'), alignment=TA_CENTER, spaceAfter=4)
    sub_style   = ParagraphStyle('hsub',   fontSize=9,  fontName='Helvetica',
                                  textColor=colors.HexColor('#64748B'), alignment=TA_CENTER, spaceAfter=4)

    elements.append(Paragraph("Rukmini Enterprises - Stock Update History", title_style))
    elements.append(Paragraph(
        f"Product: {product.name}  |  Company: {product.company_name or '—'}  |  Category: {product.category.name if product.category else '—'}",
        sub_style
    ))
    elements.append(Paragraph(
        f"Generated: {timezone.localdate().strftime('%d %b %Y')}   |   Current Stock: {product.stock_quantity} units",
        sub_style
    ))
    elements.append(Spacer(1, 0.4*cm))

    # Summary strip
    total_added   = sum(e.qty_added for e in history if e.qty_added > 0)
    total_removed = abs(sum(e.qty_added for e in history if e.qty_added < 0))
    summary_data = [[
        "Total Updates", str(history.count()),
        "Total Added", str(total_added),
        "Total Removed", str(total_removed),
        "Current Stock", str(product.stock_quantity),
    ]]
    sw = [3.46*cm] * 8
    summary_tbl = Table(summary_data, colWidths=sw)
    summary_tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), colors.HexColor('#EEF2FF')),
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,-1), 8.5),
        ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('BOX',           (0,0), (-1,-1), 0.6, colors.HexColor('#C7D2FE')),
        ('INNERGRID',     (0,0), (-1,-1), 0.4, colors.HexColor('#C7D2FE')),
        ('TOPPADDING',    (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ]))
    elements.append(summary_tbl)
    elements.append(Spacer(1, 0.4*cm))

    # Detail table
    col_headers = ["#", "Date of Stock Entry", "Qty Added / Removed", "Total Stock After Update", "Logged At"]
    col_widths_pdf = [1.2*cm, 5.5*cm, 5.5*cm, 6.0*cm, 6.5*cm]
    data = [col_headers]

    for i, entry in enumerate(history, 1):
        qty_text = f"+{entry.qty_added}" if entry.qty_added > 0 else str(entry.qty_added)
        data.append([
            str(i),
            entry.date_entered.strftime("%d %b %Y"),
            qty_text,
            f"{entry.qty_after} units",
            timezone.localtime(entry.recorded_at).strftime("%d %b %Y, %I:%M %p"),
        ])

    if len(data) == 1:
        data.append(["", "No history records found.", "", "", ""])

    tbl = Table(data, colWidths=col_widths_pdf, repeatRows=1)

    # Build row-by-row colours for added vs removed
    row_styles = [
        ('BACKGROUND',    (0,0), (-1,0),  colors.HexColor('#4F46E5')),
        ('TEXTCOLOR',     (0,0), (-1,0),  colors.white),
        ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0),  9),
        ('ALIGN',         (0,0), (-1,0),  'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,0),  7),
        ('BOTTOMPADDING', (0,0), (-1,0),  7),
        ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1), (-1,-1), 8.5),
        ('TOPPADDING',    (0,1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,1), (-1,-1), 5),
        ('ALIGN',         (0,1), (0,-1),  'CENTER'),
        ('ALIGN',         (1,1), (1,-1),  'CENTER'),
        ('ALIGN',         (2,1), (2,-1),  'CENTER'),
        ('ALIGN',         (3,1), (3,-1),  'CENTER'),
        ('ALIGN',         (4,1), (4,-1),  'CENTER'),
        ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor('#E2E8F0')),
        ('LEFTPADDING',   (0,0), (-1,-1), 5),
        ('RIGHTPADDING',  (0,0), (-1,-1), 5),
    ]
    # Colour qty column per row
    for i, entry in enumerate(history, 1):
        if entry.qty_added > 0:
            row_styles.append(('TEXTCOLOR', (2, i), (2, i), colors.HexColor('#065F46')))
            row_styles.append(('BACKGROUND', (2, i), (2, i), colors.HexColor('#D1FAE5')))
        else:
            row_styles.append(('TEXTCOLOR', (2, i), (2, i), colors.HexColor('#991B1B')))
            row_styles.append(('BACKGROUND', (2, i), (2, i), colors.HexColor('#FEE2E2')))
        # Alternating row bg
        if i % 2 == 0:
            row_styles.append(('BACKGROUND', (0, i), (1, i), colors.HexColor('#F8FAFC')))
            row_styles.append(('BACKGROUND', (3, i), (4, i), colors.HexColor('#F8FAFC')))

    tbl.setStyle(TableStyle(row_styles))
    elements.append(tbl)

    doc.build(elements)
    buffer.seek(0)
    safe_name = product.name.replace(" ", "_")[:30]
    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="stock_history_{safe_name}_{timezone.localdate()}.pdf"'
    return response


# ── All Stock History: Excel Download ────────────────────────────
@admin_required
def stock_all_history_excel(request):
    from .models import StockHistory
    history_qs = StockHistory.objects.select_related('product').order_by('product__name', '-recorded_at')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock History"

    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    group_fill  = PatternFill("solid", fgColor="EEF2FF")
    add_fill    = PatternFill("solid", fgColor="D1FAE5")
    rem_fill    = PatternFill("solid", fgColor="FEE2E2")
    bold        = Font(bold=True)
    center      = Alignment(horizontal="center", vertical="center")
    thin        = Side(style="thin", color="D1D5DB")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws.merge_cells("A1:F1")
    ws["A1"] = "Rukmini Enterprises — All Stock Update History"
    ws["A1"].font = Font(bold=True, size=14, color="1E293B")
    ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = f"Generated: {timezone.localdate().strftime('%d %b %Y')}  |  Total Records: {history_qs.count()}"
    ws["A2"].font = Font(size=10, color="64748B")
    ws["A2"].alignment = center
    ws.append([])

    # Column headers
    headers = ["#", "Product Name", "Date of Stock Entry", "Qty Added / Removed", "Total Stock After", "Logged At"]
    ws.append(headers)
    for col, cell in enumerate(ws[4], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    current_product = None
    row_num = 0
    for entry in history_qs:
        # Group header row per product
        if entry.product != current_product:
            current_product = entry.product
            ws.append([
                "", f"▶  {entry.product.name}  —  {entry.product.company_name or '—'}",
                "", "", "", ""
            ])
            grp_row = ws.max_row
            ws.merge_cells(f"B{grp_row}:F{grp_row}")
            for cell in ws[grp_row]:
                cell.fill = group_fill
                cell.font = Font(bold=True, size=10, color="4F46E5")
                cell.alignment = Alignment(vertical="center")

        row_num += 1
        qty_str = f"+{entry.qty_added}" if entry.qty_added > 0 else str(entry.qty_added)
        ws.append([
            row_num,
            entry.product.name,
            entry.date_entered.strftime("%d %b %Y"),
            qty_str,
            entry.qty_after,
            timezone.localtime(entry.recorded_at).strftime("%d %b %Y, %I:%M %p"),
        ])
        row = ws.max_row
        fill = add_fill if entry.qty_added > 0 else rem_fill
        for cell in ws[row]:
            cell.border = border
            cell.alignment = Alignment(vertical="center", horizontal="center")
        ws.cell(row=row, column=4).fill = fill

    if not history_qs.exists():
        ws.append(["", "No stock history records found.", "", "", "", ""])

    for i, w in enumerate([5, 28, 22, 22, 22, 28], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="all_stock_history_{timezone.localdate()}.xlsx"'
    wb.save(response)
    return response


# ── All Stock History: PDF Download ────────────────────────────
@admin_required
def stock_all_history_pdf(request):
    from .models import StockHistory
    history_qs = StockHistory.objects.select_related('product').order_by('product__name', '-recorded_at')

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    elements = []

    title_style = ParagraphStyle('ahtitle', fontSize=15, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#1E293B'), alignment=TA_CENTER, spaceAfter=4)
    sub_style   = ParagraphStyle('ahsub', fontSize=9, fontName='Helvetica',
                                  textColor=colors.HexColor('#64748B'), alignment=TA_CENTER, spaceAfter=12)
    group_style = ParagraphStyle('ahgrp', fontSize=9, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#4F46E5'))

    elements.append(Paragraph("Rukmini Enterprises - All Stock Update History", title_style))
    elements.append(Paragraph(
        f"Generated: {timezone.localdate().strftime('%d %b %Y')}   |   Total Records: {history_qs.count()}",
        sub_style
    ))

    # Summary strip
    total_added   = sum(e.qty_added for e in history_qs if e.qty_added > 0)
    total_removed = abs(sum(e.qty_added for e in history_qs if e.qty_added < 0))
    summary_data = [[
        "Total Records", str(history_qs.count()),
        "Total Added", str(total_added),
        "Total Removed", str(total_removed),
        "Products Tracked", str(history_qs.values('product').distinct().count()),
    ]]
    sw = [3.46*cm] * 8
    summary_tbl = Table(summary_data, colWidths=sw)
    summary_tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), colors.HexColor('#EEF2FF')),
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,-1), 8.5),
        ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('BOX',           (0,0), (-1,-1), 0.6, colors.HexColor('#C7D2FE')),
        ('INNERGRID',     (0,0), (-1,-1), 0.4, colors.HexColor('#C7D2FE')),
        ('TOPPADDING',    (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ]))
    elements.append(summary_tbl)
    elements.append(Spacer(1, 0.4*cm))

    # Detail table
    col_headers = ["#", "Product Name", "Date of Entry", "Qty Added / Removed", "Stock After", "Logged At"]
    col_widths_pdf = [1.0*cm, 6.5*cm, 4.5*cm, 4.5*cm, 3.5*cm, 5.7*cm]
    data = [col_headers]
    row_styles = [
        ('BACKGROUND',    (0,0), (-1,0),  colors.HexColor('#4F46E5')),
        ('TEXTCOLOR',     (0,0), (-1,0),  colors.white),
        ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0),  9),
        ('ALIGN',         (0,0), (-1,0),  'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,0),  6),
        ('BOTTOMPADDING', (0,0), (-1,0),  6),
        ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1), (-1,-1), 8),
        ('TOPPADDING',    (0,1), (-1,-1), 4),
        ('BOTTOMPADDING', (0,1), (-1,-1), 4),
        ('ALIGN',         (0,1), (0,-1),  'CENTER'),
        ('ALIGN',         (1,1), (1,-1),  'LEFT'),
        ('ALIGN',         (2,1), (4,-1),  'CENTER'),
        ('ALIGN',         (5,1), (5,-1),  'CENTER'),
        ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor('#E2E8F0')),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
    ]

    current_product = None
    row_num = 0
    data_row_index = 1  # header is index 0

    for entry in history_qs:
        # Group header row per product
        if entry.product != current_product:
            current_product = entry.product
            data.append([f"▶  {entry.product.name}  —  {entry.product.company_name or '—'}", "", "", "", "", ""])
            row_styles.append(('BACKGROUND', (0, data_row_index), (-1, data_row_index), colors.HexColor('#EEF2FF')))
            row_styles.append(('TEXTCOLOR',  (0, data_row_index), (-1, data_row_index), colors.HexColor('#4F46E5')))
            row_styles.append(('FONTNAME',   (0, data_row_index), (-1, data_row_index), 'Helvetica-Bold'))
            row_styles.append(('SPAN',       (0, data_row_index), (-1, data_row_index)))
            data_row_index += 1

        row_num += 1
        qty_str = f"+{entry.qty_added}" if entry.qty_added > 0 else str(entry.qty_added)
        data.append([
            str(row_num),
            entry.product.name[:28],
            entry.date_entered.strftime("%d %b %Y"),
            qty_str,
            f"{entry.qty_after} units",
            timezone.localtime(entry.recorded_at).strftime("%d %b %y %I:%M %p"),
        ])
        if entry.qty_added > 0:
            row_styles.append(('TEXTCOLOR',  (3, data_row_index), (3, data_row_index), colors.HexColor('#065F46')))
            row_styles.append(('BACKGROUND', (3, data_row_index), (3, data_row_index), colors.HexColor('#D1FAE5')))
        else:
            row_styles.append(('TEXTCOLOR',  (3, data_row_index), (3, data_row_index), colors.HexColor('#991B1B')))
            row_styles.append(('BACKGROUND', (3, data_row_index), (3, data_row_index), colors.HexColor('#FEE2E2')))
        if row_num % 2 == 0:
            row_styles.append(('BACKGROUND', (0, data_row_index), (2, data_row_index), colors.HexColor('#F8FAFC')))
            row_styles.append(('BACKGROUND', (4, data_row_index), (5, data_row_index), colors.HexColor('#F8FAFC')))
        data_row_index += 1

    if len(data) == 1:
        data.append(["", "No stock history records found.", "", "", "", ""])

    tbl = Table(data, colWidths=col_widths_pdf, repeatRows=1)
    tbl.setStyle(TableStyle(row_styles))
    elements.append(tbl)

    doc.build(elements)
    buffer.seek(0)
    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="all_stock_history_{timezone.localdate()}.pdf"'
    return response
