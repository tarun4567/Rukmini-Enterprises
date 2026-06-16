from django.contrib import admin
from .models import Category, Product, Bill, BillItem, StockHistory, Expense

@admin.register(Category)
class CategoryAdmin(admin.ModelAdmin):
    list_display = ('name', 'slug')
    prepopulated_fields = {'slug': ('name',)}
    search_fields = ('name',)

@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ('name', 'category', 'purchase_price', 'selling_price', 'stock_quantity', 'min_stock_level', 'is_low_stock')
    list_filter = ('category', 'stock_quantity')
    search_fields = ('name', 'company_name')
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


@admin.register(StockHistory)
class StockHistoryAdmin(admin.ModelAdmin):
    list_display  = ('product', 'date_entered', 'qty_added', 'qty_after', 'recorded_at')
    list_filter   = ('date_entered', 'product')
    search_fields = ('product__name',)
    readonly_fields = ('recorded_at',)


@admin.register(Expense)
class ExpenseAdmin(admin.ModelAdmin):
    list_display  = ('company_name', 'amount', 'date_paid', 'recorded_by', 'created_at')
    list_filter   = ('date_paid', 'company_name', 'recorded_by')
    search_fields = ('company_name', 'description')
    readonly_fields = ('created_at',)
