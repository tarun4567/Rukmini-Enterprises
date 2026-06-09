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

from .models import Category, Product, Customer, Order, OrderItem, Bill
from .forms import ProductForm, StockForm

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
        
        total = quantity * rate
        amount_to_be_given = amount_given - total
        
        if product.stock_quantity < quantity:
            messages.warning(
                request,
                f"Warning: Quantity ordered ({quantity}) exceeds stock level ({product.stock_quantity}) for {product.name}."
            )
            
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
        
        messages.success(request, f"Bill #{bill.id} for {cust_name} created successfully! Change to return: ${amount_to_be_given:.2f}")
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
            
    daily_bills = Bill.objects.filter(created_at__date=query_date).order_by('-created_at')
    
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
                        messages.success(request, f"Payment of ${clear_amount:.2f} received. Bill #{bill.id} is now fully cleared!")
                    else:
                        messages.success(request, f"Payment of ${clear_amount:.2f} received for Bill #{bill.id}. Remaining due: ${abs(bill.amount_to_be_given):.2f}")
                else:
                    messages.error(request, "Payment amount must be greater than zero.")
            except ValueError:
                messages.error(request, "Invalid payment amount format.")
        else:
            messages.error(request, "No payment amount was specified.")
            
        bill_date_str = bill.created_at.strftime('%Y-%m-%d')
        return redirect(f"/billing/records/?date={bill_date_str}")
    return redirect('billing_records')


# 5. Dashboard View (Admin Only)
@admin_required
def dashboard_view(request):
    total_products = Product.objects.count()
    low_stock_products = Product.objects.filter(stock_quantity__lte=F('min_stock_level'))
    low_stock_count = low_stock_products.count()
    total_customers = Customer.objects.count()
    
    revenue_orders = Order.objects.filter(status__in=['Paid', 'Shipped'])
    orders_rev = sum(order.total_price for order in revenue_orders)
    bills_rev = float(Bill.objects.aggregate(total=Sum('total'))['total'] or 0.00)
    total_revenue = float(orders_rev) + bills_rev
    
    today = timezone.now()
    sales_labels = []
    sales_values = []
    for i in range(5, -1, -1):
        month_date = today - datetime.timedelta(days=i*30)
        month_name = month_date.strftime("%b %Y")
        sales_labels.append(month_name)
        
        month_orders = Order.objects.filter(
            status__in=['Paid', 'Shipped'],
            order_date__year=month_date.year,
            order_date__month=month_date.month
        )
        month_order_total = sum(order.total_price for order in month_orders)
        
        month_bills = Bill.objects.filter(
            created_at__year=month_date.year,
            created_at__month=month_date.month
        )
        month_bill_total = sum(bill.total for bill in month_bills)
        
        month_total = float(month_order_total) + float(month_bill_total)
        sales_values.append(month_total)
        
    category_labels = []
    category_values = []
    for category in Category.objects.all():
        val = sum(p.stock_quantity * p.selling_price for p in category.products.all())
        if val > 0:
            category_labels.append(category.name)
            category_values.append(float(val))

    recent_orders = Order.objects.order_by('-order_date')[:5]
    
    context = {
        'total_products': total_products,
        'low_stock_count': low_stock_count,
        'low_stock_products': low_stock_products[:5],
        'total_customers': total_customers,
        'total_revenue': total_revenue,
        'recent_orders': recent_orders,
        'sales_labels': json.dumps(sales_labels),
        'sales_values': json.dumps(sales_values),
        'category_labels': json.dumps(category_labels),
        'category_values': json.dumps(category_values),
    }
    return render(request, 'dashboard.html', context)


# 6. Inventory View & CRUD (Admin Only)
@admin_required
def inventory_view(request):
    categories = Category.objects.all()
    active_category = request.GET.get('category')
    
    products = Product.objects.all()
    if active_category:
        products = products.filter(category__slug=active_category)
        
    context = {
        'products': products,
        'categories': categories,
        'active_category': active_category,
    }
    return render(request, 'inventory.html', context)


@admin_required
def product_add_view(request):
    categories = Category.objects.all()
    if request.method == 'POST':
        form = ProductForm(request.POST, request.FILES)
        if form.is_valid():
            product = form.save()
            messages.success(request, f"Product '{product.name}' was added successfully.")
            return redirect('inventory')
    else:
        form = ProductForm()
    return render(request, 'product_form.html', {'form': form, 'categories': categories})


@admin_required
def product_edit_view(request, pk):
    product = get_object_or_404(Product, pk=pk)
    categories = Category.objects.all()
    if request.method == 'POST':
        form = ProductForm(request.POST, request.FILES, instance=product)
        if form.is_valid():
            product = form.save()
            messages.success(request, f"Product '{product.name}' was updated successfully.")
            return redirect('inventory')
    else:
        form = ProductForm(instance=product)
    return render(request, 'product_form.html', {'form': form, 'categories': categories})


@admin_required
def product_delete_view(request, pk):
    product = get_object_or_404(Product, pk=pk)
    if request.method == 'POST':
        name = product.name
        product.delete()
        messages.warning(request, f"Product '{name}' was deleted from inventory.")
    return redirect('inventory')


# 7. Customers CRM (Admin Only)
@admin_required
def customers_view(request):
    customers = Customer.objects.order_by('-date_created')
    return render(request, 'customers.html', {'customers': customers})


@admin_required
def customer_add_view(request):
    if request.method == 'POST':
        name = request.POST.get('name')
        email = request.POST.get('email')
        phone = request.POST.get('phone')
        company_name = request.POST.get('company_name')
        address = request.POST.get('address')
        
        if name:
            customer = Customer.objects.create(
                name=name, email=email, phone=phone,
                company_name=company_name, address=address
            )
            messages.success(request, f"Customer '{customer.name}' registered successfully.")
        else:
            messages.error(request, "Failed to register customer: Name is required.")
    return redirect('customers')


# 8. Orders View & CRUD (Admin Only)
@admin_required
def orders_view(request):
    status_choices = Order.STATUS_CHOICES
    active_status = request.GET.get('status')
    
    orders = Order.objects.order_by('-order_date')
    if active_status:
        orders = orders.filter(status=active_status)
        
    products = Product.objects.all()
    customers = Customer.objects.all()
    
    context = {
        'orders': orders,
        'products': products,
        'customers': customers,
        'status_choices': status_choices,
        'active_status': active_status,
    }
    return render(request, 'orders.html', context)


@admin_required
def order_add_view(request):
    if request.method == 'POST':
        customer_id = request.POST.get('customer')
        product_id = request.POST.get('product')
        quantity = int(request.POST.get('quantity', 1))
        status = request.POST.get('status', 'Pending')
        notes = request.POST.get('notes')
        
        customer = get_object_or_404(Customer, id=customer_id)
        product = get_object_or_404(Product, id=product_id)
        
        if product.stock_quantity < quantity:
            messages.warning(
                request,
                f"Warning: Ordered quantity ({quantity}) exceeds current in-stock level ({product.stock_quantity}) for {product.name}."
            )
        
        product.stock_quantity -= quantity
        product.save()
        
        order = Order.objects.create(customer=customer, status=status, notes=notes)
        OrderItem.objects.create(
            order=order,
            product=product,
            quantity=quantity,
            price=product.selling_price
        )
        
        messages.success(request, f"Sales Order #{order.id} for {customer.name} created successfully.")
    return redirect('orders')


@admin_required
def order_update_status_view(request, pk):
    order = get_object_or_404(Order, pk=pk)
    if request.method == 'POST':
        old_status = order.status
        new_status = request.POST.get('status')
        if new_status in dict(Order.STATUS_CHOICES):
            order.status = new_status
            order.save()
            messages.success(request, f"Order #{order.id} status updated from {old_status} to {new_status}.")
    return redirect('orders')


@admin_required
def order_delete_view(request, pk):
    order = get_object_or_404(Order, pk=pk)
    if request.method == 'POST':
        order_id = order.id
        for item in order.items.all():
            if item.product:
                item.product.stock_quantity += item.quantity
                item.product.save()
        order.delete()
        messages.warning(request, f"Order #{order_id} has been cancelled and deleted. Stock levels restored.")
    return redirect('orders')


# 9. Stock & Stock 1 Views (Admin Only)
@admin_required
def stock_add_view(request):
    categories = Category.objects.all()
    products = Product.objects.all().order_by('-id')
    
    if request.method == 'POST':
        form = StockForm(request.POST, request.FILES)
        if form.is_valid():
            product = form.save(commit=False)
            product.initial_quantity = product.stock_quantity
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
    categories = Category.objects.all()
    active_category = request.GET.get('category')
    
    products = Product.objects.all().order_by('-id')
    if active_category:
        products = products.filter(category__slug=active_category)
        
    context = {
        'products': products,
        'categories': categories,
        'active_category': active_category,
    }
    return render(request, 'stock1.html', context)


@admin_required
def stock_edit_view(request, pk):
    product = get_object_or_404(Product, pk=pk)
    categories = Category.objects.all()
    products = Product.objects.all().order_by('-id')
    
    if request.method == 'POST':
        old_qty = product.stock_quantity
        form = StockForm(request.POST, request.FILES, instance=product)
        if form.is_valid():
            product = form.save(commit=False)
            qty_diff = product.stock_quantity - old_qty
            if qty_diff != 0:
                product.initial_quantity = max(0, (product.initial_quantity or 0) + qty_diff)
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
