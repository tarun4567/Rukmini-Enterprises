from django.contrib import admin
from .models import Category, Product, Bill

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




@admin.register(Bill)
class BillAdmin(admin.ModelAdmin):
    list_display = ('id', 'customer_name', 'product', 'quantity', 'rate', 'total', 'created_at')
    list_filter = ('created_at', 'product')
    search_fields = ('customer_name', 'product__name')
