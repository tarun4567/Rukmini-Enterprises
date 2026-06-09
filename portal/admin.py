from django.contrib import admin
from .models import Category, Product, Bill, BillItem

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


class BillItemInline(admin.TabularInline):
    model = BillItem
    extra = 0
    readonly_fields = ('product', 'quantity', 'rate', 'total')


@admin.register(Bill)
class BillAdmin(admin.ModelAdmin):
    list_display  = ('id', 'customer_name', 'grand_total', 'amount_given', 'amount_to_be_given', 'created_at')
    list_filter   = ('created_at',)
    search_fields = ('customer_name',)
    inlines       = [BillItemInline]
