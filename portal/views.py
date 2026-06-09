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

from .models import Category, Product, Bill, BillItem
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
        cust_address = request.POST.get('customer_address', '').strip()
        amount_given = float(request.POST.get('amount_given', 0.00))

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
    import calendar
    total_products = Product.objects.count()
    low_stock_products = Product.objects.filter(stock_quantity__lte=F('min_stock_level'))
    low_stock_count = low_stock_products.count()
    
    total_revenue = float(Bill.objects.aggregate(total=Sum('grand_total'))['total'] or 0.00)
    
    # Year & Month dynamic filters
    current_year = timezone.now().year
    
    # Collect available years from DB
    raw_years = (
        [current_year] + 
        [y.year for y in Bill.objects.dates('created_at', 'year') if hasattr(y, 'year')]
    )
    available_years = sorted(list(set(raw_years)), reverse=True)
    if len(available_years) < 2:
        available_years = [current_year, current_year - 1]

    months_list = [
        (1, "January"), (2, "February"), (3, "March"), (4, "April"),
        (5, "May"), (6, "June"), (7, "July"), (8, "August"),
        (9, "September"), (10, "October"), (11, "November"), (12, "December")
    ]
    
    selected_year_str = request.GET.get('year')
    selected_month_str = request.GET.get('month', 'all')
    
    try:
        selected_year = int(selected_year_str) if selected_year_str else current_year
    except ValueError:
        selected_year = current_year
        
    selected_month = selected_month_str
    selected_month_name = None
    
    sales_labels = []
    sales_values = []
    
    if selected_month == 'all':
        # Monthly trend for the selected year
        sales_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        for month in range(1, 13):
            month_bills = Bill.objects.filter(
                created_at__year=selected_year,
                created_at__month=month
            )
            month_total = float(month_bills.aggregate(total=Sum('grand_total'))['total'] or 0.00)
            sales_values.append(month_total)
    else:
        try:
            m_int = int(selected_month)
            if 1 <= m_int <= 12:
                selected_month_name = dict(months_list).get(m_int)
                days_in_month = calendar.monthrange(selected_year, m_int)[1]
                sales_labels = [str(d) for d in range(1, days_in_month + 1)]
                for day in range(1, days_in_month + 1):
                    local_date = datetime.date(selected_year, m_int, day)
                    start_of_day = timezone.make_aware(datetime.datetime.combine(local_date, datetime.time.min))
                    end_of_day = timezone.make_aware(datetime.datetime.combine(local_date, datetime.time.max))
                    
                    day_bills = Bill.objects.filter(
                        created_at__range=(start_of_day, end_of_day)
                    )
                    day_total = float(day_bills.aggregate(total=Sum('grand_total'))['total'] or 0.00)
                    sales_values.append(day_total)
            else:
                selected_month = 'all'
        except ValueError:
            selected_month = 'all'

        # Fallback if invalid month fell back to 'all'
        if selected_month == 'all':
            sales_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            for month in range(1, 13):
                month_bills = Bill.objects.filter(
                    created_at__year=selected_year,
                    created_at__month=month
                )
                month_total = float(month_bills.aggregate(total=Sum('grand_total'))['total'] or 0.00)
                sales_values.append(month_total)
                
    category_labels = []
    category_values = []
    for category in Category.objects.all():
        val = sum(p.stock_quantity * p.selling_price for p in category.products.all())
        if val > 0:
            category_labels.append(category.name)
            category_values.append(float(val))
    
    context = {
        'total_products': total_products,
        'low_stock_count': low_stock_count,
        'low_stock_products': low_stock_products[:5],
        'total_revenue': total_revenue,
        'sales_labels': json.dumps(sales_labels),
        'sales_values': json.dumps(sales_values),
        'category_labels': json.dumps(category_labels),
        'category_values': json.dumps(category_values),
        'available_years': available_years,
        'months_list': months_list,
        'selected_year': selected_year,
        'selected_month': selected_month,
        'selected_month_name': selected_month_name,
    }
    return render(request, 'dashboard.html', context)


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
            messages.success(request, f"Stock item '{product.name}' was updated successfully.")
            return redirect('stock_add')
    else:
        form = StockForm(instance=product)
        
    context = {
        'form': form,
        'products': products,
        'categories': categories,
        'product': product,
        'is_edit': True,
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
