from django.db import models
from django.utils.text import slugify

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
    sku = models.CharField(max_length=50, unique=True, verbose_name="SKU", blank=True)
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
    stock_entered = models.CharField(max_length=100, blank=True, null=True)
    initial_quantity = models.IntegerField(default=0)

    def save(self, *args, **kwargs):
        # 1. Assign Default Category if not set
        if not self.category:
            default_cat, _ = Category.objects.get_or_create(
                name="General",
                defaults={"description": "General stock catalog item category."}
            )
            self.category = default_cat

        # 2. Auto-generate unique SKU if not set
        if not self.sku:
            import uuid
            company_slug = slugify(self.company_name or "GEN")[:4].upper()
            product_slug = slugify(self.name or "PROD")[:4].upper()
            unique_id = str(uuid.uuid4())[:4].upper()
            self.sku = f"STK-{company_slug}-{product_slug}-{unique_id}"

        # 3. Default purchase price to selling price (Amt) if 0 or not set
        if not self.purchase_price or self.purchase_price == 0:
            self.purchase_price = self.selling_price

        # 4. Set initial_quantity on first save or if it is currently 0/unset
        if not self.pk or not self.initial_quantity or self.initial_quantity == 0:
            self.initial_quantity = self.stock_quantity

        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name} ({self.sku})"

    @property
    def is_low_stock(self):
        return self.stock_quantity <= self.min_stock_level

    @property
    def total_value(self):
        return self.stock_quantity * self.selling_price

    @property
    def total_purchase_value(self):
        return self.stock_quantity * self.purchase_price


class Customer(models.Model):
    name = models.CharField(max_length=150)
    email = models.EmailField(blank=True, null=True)
    phone = models.CharField(max_length=20, blank=True, null=True)
    company_name = models.CharField(max_length=150, blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    date_created = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        if self.company_name:
            return f"{self.name} - {self.company_name}"
        return self.name

    @property
    def total_orders(self):
        return self.orders.count()

    @property
    def total_spent(self):
        total = sum(order.total_price for order in self.orders.all())
        return total


class Order(models.Model):
    STATUS_CHOICES = [
        ('Draft', 'Draft'),
        ('Pending', 'Pending'),
        ('Paid', 'Paid'),
        ('Shipped', 'Shipped'),
        ('Cancelled', 'Cancelled'),
    ]

    customer = models.ForeignKey(Customer, on_delete=models.CASCADE, related_name="orders")
    order_date = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='Pending')
    notes = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"Order #{self.id} - {self.customer.name}"

    @property
    def total_price(self):
        return sum(item.total_price for item in self.items.all())

    @property
    def total_items(self):
        return sum(item.quantity for item in self.items.all())


class OrderItem(models.Model):
    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name="items")
    product = models.ForeignKey(Product, on_delete=models.SET_NULL, null=True)
    quantity = models.IntegerField(default=1)
    price = models.DecimalField(max_digits=10, decimal_places=2, help_text="Price of the item at the time of ordering")

    def __str__(self):
        return f"{self.quantity} x {self.product.name if self.product else 'Deleted Product'} in Order #{self.order.id}"

    @property
    def total_price(self):
        return self.quantity * self.price


class Bill(models.Model):
    customer_name = models.CharField(max_length=150, verbose_name="Customer Name")
    customer_address = models.TextField(blank=True, null=True, verbose_name="Customer Address")
    product = models.ForeignKey(Product, on_delete=models.SET_NULL, null=True)
    quantity = models.IntegerField(default=1)
    rate = models.DecimalField(max_digits=10, decimal_places=2)
    total = models.DecimalField(max_digits=10, decimal_places=2)
    amount_given = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Amount Given")
    amount_to_be_given = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Amount to be Given")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Bill #{self.id} - {self.customer_name}"

    @property
    def abs_amount_to_be_given(self):
        return abs(self.amount_to_be_given)

