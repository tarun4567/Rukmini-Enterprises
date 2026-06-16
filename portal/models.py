from django.db import models
from django.utils.text import slugify
from django.utils import timezone

class Category(models.Model):
    name = models.CharField(max_length=100, unique=True)
    slug = models.SlugField(max_length=100, unique=True, blank=True)
    description = models.TextField(blank=True, null=True)

    class Meta:
        verbose_name_plural = "Categories"

    def save(self, *args, **kwargs):
        if not self.slug:
            self.slug = slugify(self.name)
        super().save(*args, **kwargs)

    def __str__(self):
        return self.name


class Product(models.Model):
    name = models.CharField(max_length=200)
    category = models.ForeignKey(Category, on_delete=models.CASCADE, related_name="products", blank=True, null=True)
    description = models.TextField(blank=True, null=True)
    purchase_price = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, blank=True)
    selling_price = models.DecimalField(max_digits=10, decimal_places=2)
    stock_quantity = models.IntegerField(default=0)
    min_stock_level = models.IntegerField(default=5)
    image = models.ImageField(upload_to="products/", blank=True, null=True)
    
    # New fields
    photo = models.ImageField(upload_to="photos/", blank=True, null=True)
    company_name = models.CharField(max_length=200, blank=True, null=True)
    stock_entered = models.DateField(blank=True, null=True)
    initial_quantity = models.IntegerField(default=0)

    def save(self, *args, **kwargs):
        # 1. Assign Default Category if not set
        if not self.category:
            default_cat, _ = Category.objects.get_or_create(
                name="General",
                defaults={"description": "General stock catalog item category."}
            )
            self.category = default_cat

        # 2. Default purchase price to selling price (Amt) if 0 or not set
        if not self.purchase_price or self.purchase_price == 0:
            self.purchase_price = self.selling_price

        # 4. Sync initial_quantity and stock_quantity
        if not self.initial_quantity or self.initial_quantity == 0:
            self.initial_quantity = self.stock_quantity
        elif not self.stock_quantity or self.stock_quantity == 0:
            self.stock_quantity = self.initial_quantity

        super().save(*args, **kwargs)

    def __str__(self):
        return self.name

    @property
    def is_low_stock(self):
        return self.stock_quantity <= self.min_stock_level

    @property
    def total_value(self):
        return self.stock_quantity * self.selling_price

    @property
    def total_purchase_value(self):
        return self.stock_quantity * self.purchase_price


class Bill(models.Model):
    customer_name    = models.CharField(max_length=150, verbose_name="Customer Name")
    customer_phone   = models.CharField(max_length=10, blank=True, null=True, verbose_name="Phone Number")
    customer_address = models.TextField(blank=True, null=True, verbose_name="Customer Address")
    grand_total      = models.DecimalField(max_digits=12, decimal_places=2, default=0.00)
    amount_given     = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Amount Given")
    amount_to_be_given = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Amount to be Given")
    created_at       = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Bill #{self.id} - {self.customer_name}"

    @property
    def abs_amount_to_be_given(self):
        return abs(self.amount_to_be_given)

    # ── Legacy compatibility properties (for views/templates that still reference these) ──
    @property
    def total(self):
        return self.grand_total

    @property
    def items_list(self):
        return self.items.select_related('product').all()


class BillItem(models.Model):
    bill     = models.ForeignKey(Bill, on_delete=models.CASCADE, related_name='items')
    product  = models.ForeignKey(Product, on_delete=models.SET_NULL, null=True)
    quantity = models.IntegerField(default=1)
    rate     = models.DecimalField(max_digits=10, decimal_places=2)
    total    = models.DecimalField(max_digits=10, decimal_places=2)

    def __str__(self):
        return f"Item: {self.product.name if self.product else 'Deleted'} x {self.quantity}"


class StockHistory(models.Model):
    """Logs every stock quantity update made by the admin."""
    product       = models.ForeignKey(Product, on_delete=models.CASCADE, related_name='stock_history')
    date_entered  = models.DateField(verbose_name="Stock Date")
    qty_added     = models.IntegerField(verbose_name="Qty Added")
    qty_after     = models.IntegerField(verbose_name="Total Qty After Update")
    recorded_at   = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-recorded_at']

    def __str__(self):
        return f"{self.product.name} | +{self.qty_added} on {self.date_entered}"


class Expense(models.Model):
    company_name = models.CharField(max_length=200, verbose_name="Company/Vendor Name")
    amount       = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Amount Paid")
    description  = models.TextField(blank=True, null=True, verbose_name="Notes / Description")
    date_paid    = models.DateField(default=timezone.localdate, verbose_name="Payment Date")
    recorded_by  = models.ForeignKey('auth.User', on_delete=models.SET_NULL, null=True, related_name='expenses')
    created_at   = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-date_paid', '-created_at']

    def __str__(self):
        return f"Expense of ₹{self.amount} to {self.company_name} on {self.date_paid}"

