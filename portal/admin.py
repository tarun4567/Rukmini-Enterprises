from django.contrib import admin
from .models import Category, Product, Customer, Order, OrderItem, Bill

@admin.register(Category)
class CategoryAdmin(admin.ModelAdmin):
    list_display = ('name', 'slug')
    prepopulated_fields = {'slug': ('name',)}
    search_fields = ('name',)

@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ('name', 'sku', 'category', 'purchase_price', 'selling_price', 'stock_quantity', 'min_stock_level', 'is_low_stock')
    list_filter = ('category', 'stock_quantity')
    search_fields = ('name', 'sku')
    list_editable = ('purchase_price', 'selling_price', 'stock_quantity', 'min_stock_level')

class OrderItemInline(admin.TabularInline):
    model = OrderItem
    extra = 1

@admin.register(Customer)
class CustomerAdmin(admin.ModelAdmin):
    list_display = ('name', 'email', 'phone', 'company_name', 'date_created')
    search_fields = ('name', 'company_name', 'email')

@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display = ('id', 'customer', 'order_date', 'status', 'total_price')
    list_filter = ('status', 'order_date')
    search_fields = ('customer__name', 'id')
    inlines = [OrderItemInline]


@admin.register(Bill)
class BillAdmin(admin.ModelAdmin):
    list_display = ('id', 'customer_name', 'product', 'quantity', 'rate', 'total', 'created_at')
    list_filter = ('created_at', 'product')
    search_fields = ('customer_name', 'product__name')
