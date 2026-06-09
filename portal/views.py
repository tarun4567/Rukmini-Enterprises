from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import AuthenticationForm
from django.contrib import messages
from django.db.models import F, Sum
from django.utils import timezone
from functools import wraps
import datetime
import json

from .models import Category, Product, Bill
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
        cust_name = request.POST.get('customer_name')
        cust_address = request.POST.get('customer_address')
        product_id = request.POST.get('product')
        quantity = int(request.POST.get('quantity', 1))
        rate = float(request.POST.get('rate', 0.00))
        amount_given = float(request.POST.get('amount_given', 0.00))
        
        product = get_object_or_404(Product, id=product_id)
        
        if product.stock_quantity < quantity:
            messages.error(
                request,
                f"Error: Cannot generate bill. Requested quantity ({quantity}) exceeds available stock ({product.stock_quantity}) for {product.name}."
            )
            context = {
                'products': Product.objects.all(),
                'cust_name': cust_name,
                'cust_address': cust_address,
                'selected_product_id': product.id,
                'quantity': quantity,
                'rate': rate,
                'amount_given': amount_given,
            }
            return render(request, 'billing.html', context)
            
        total = quantity * rate
        amount_to_be_given = amount_given - total
        
        product.stock_quantity -= quantity
        product.save()
        
        bill = Bill.objects.create(
            customer_name=cust_name,
            customer_address=cust_address,
            product=product,
            quantity=quantity,
            rate=rate,
            total=total,
            amount_given=amount_given,
            amount_to_be_given=amount_to_be_given
        )
        
        messages.success(request, f"Bill #{bill.id} for {cust_name} created successfully! Change to return: ₹{amount_to_be_given:.2f}")
        return redirect('billing_records')

    context = {
        'products': products,
    }
    return render(request, 'billing.html', context)


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
    
    total_revenue = float(Bill.objects.aggregate(total=Sum('total'))['total'] or 0.00)
    
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
            month_total = float(month_bills.aggregate(total=Sum('total'))['total'] or 0.00)
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
                    day_total = float(day_bills.aggregate(total=Sum('total'))['total'] or 0.00)
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
                month_total = float(month_bills.aggregate(total=Sum('total'))['total'] or 0.00)
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
    total_revenue = float(bills.aggregate(total=Sum('total'))['total'] or 0.0)
    total_sales = bills.count()
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
                day_map[local_bill_date] += float(bill.total)
                
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
                month_map[m_key] += float(bill.total)
                
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
        'avg_order_value': avg_order_value,
        'total_dues': total_dues,
        'chart_labels': json.dumps(chart_labels),
        'chart_values': json.dumps(chart_values),
    }
    return render(request, 'revenue_report.html', context)
