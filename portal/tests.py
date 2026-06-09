from django.test import TestCase, Client
from django.contrib.auth.models import User
from django.urls import reverse
from django.utils import timezone
import datetime

from .models import Category, Product, Customer, Order, OrderItem, Bill

class ERPTests(TestCase):
    def setUp(self):
        # Create an admin user (staff)
        self.admin_user = User.objects.create_user(username="testadmin", password="testpassword", is_staff=True)
        
        # Create a regular user (not staff)
        self.regular_user = User.objects.create_user(username="testuser", password="testpassword", is_staff=False)
        
        # Create category
        self.category = Category.objects.create(name="Steel Profiles")
        
        # Create product
        self.product = Product.objects.create(
            name="Steel Angle Rod",
            sku="STL-ANG-01",
            category=self.category,
            purchase_price=20.00,
            selling_price=45.00,
            stock_quantity=10,
            min_stock_level=5
        )
        
        # Create customer for CRM orders
        self.customer = Customer.objects.create(
            name="Jane Doe",
            email="jane@example.com",
            phone="1234567890"
        )
        
        self.client = Client()

    def test_product_properties(self):
        """Test Product helper properties (low stock, value)"""
        product = self.product
        self.assertEqual(product.total_value, 450.00)
        self.assertFalse(product.is_low_stock)
        
        # Drop stock below limit
        product.stock_quantity = 3
        product.save()
        self.assertTrue(product.is_low_stock)

    def test_auth_protection(self):
        """Verify unauthenticated requests are redirected to login"""
        self.client.logout()
        response = self.client.get(reverse('dashboard'))
        self.assertEqual(response.status_code, 302) # Redirect to login
        self.assertIn('/login/', response.url)

    def test_admin_dashboard_access(self):
        """Verify admin user can access dashboard, but redirected from billing/records"""
        self.client.login(username="testadmin", password="testpassword")
        response = self.client.get(reverse('dashboard'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Welcome, testadmin")
        
        # Try to access billing -> should redirect to dashboard
        response = self.client.get(reverse('billing'))
        self.assertRedirects(response, reverse('dashboard'))
        
        # Try to access billing records -> should redirect to dashboard
        response = self.client.get(reverse('billing_records'))
        self.assertRedirects(response, reverse('dashboard'))

    def test_regular_user_access_control(self):
        """Verify regular user can access billing/records/stocks but is redirected from dashboard/inventory"""
        self.client.login(username="testuser", password="testpassword")
        
        # Access dashboard -> should redirect to billing (standard user homepage)
        response = self.client.get(reverse('dashboard'))
        self.assertRedirects(response, reverse('billing'))
        
        # Access inventory management -> should redirect to billing
        response = self.client.get(reverse('inventory'))
        self.assertRedirects(response, reverse('billing'))
        
        # Access billing page -> should load successfully (HTTP 200)
        response = self.client.get(reverse('billing'))
        self.assertEqual(response.status_code, 200)
        
        # Access billing records page -> should load successfully (HTTP 200)
        response = self.client.get(reverse('billing_records'))
        self.assertEqual(response.status_code, 200)
        
        # Access stocks lookup page -> should load successfully (HTTP 200)
        response = self.client.get(reverse('stocks'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Steel Angle Rod")

    def test_billing_creation_and_stock_deduction(self):
        """Verify Day Billing form submission logs transactions and redirects to records (run by standard user)"""
        self.client.login(username="testuser", password="testpassword")
        
        response = self.client.post(reverse('billing'), {
            'customer_name': 'Walking Retail Client',
            'customer_address': 'Counter Sale',
            'product': self.product.id,
            'quantity': 3,
            'rate': 45.00,
            'amount_given': 150.00
        })
        
        # Verify redirect to billing records
        self.assertRedirects(response, reverse('billing_records'))
        
        # Verify Bill object created
        self.assertEqual(Bill.objects.count(), 1)
        bill = Bill.objects.first()
        self.assertEqual(bill.customer_name, "Walking Retail Client")
        self.assertEqual(bill.quantity, 3)
        self.assertEqual(bill.rate, 45.00)
        self.assertEqual(bill.total, 135.00) # 3 * 45
        self.assertEqual(bill.amount_given, 150.00)
        self.assertEqual(bill.amount_to_be_given, 15.00) # 150 - 135
        
        # Verify stock levels updated
        self.product.refresh_from_db()
        self.assertEqual(self.product.stock_quantity, 7) # 10 initial - 3 sold = 7

    def test_billing_records_date_search(self):
        """Verify querying a specific date lists correct bills and filters others"""
        self.client.login(username="testuser", password="testpassword")
        
        # Bill A created today
        bill_today = Bill.objects.create(
            customer_name="Today Customer", product=self.product,
            quantity=1, rate=45.00, total=45.00, amount_given=45.00, amount_to_be_given=0.00
        )
        
        # Bill B created in the past
        bill_past = Bill.objects.create(
            customer_name="Past Customer", product=self.product,
            quantity=1, rate=45.00, total=45.00, amount_given=45.00, amount_to_be_given=0.00
        )
        past_date = timezone.now() - datetime.timedelta(days=3)
        Bill.objects.filter(pk=bill_past.pk).update(created_at=past_date)
        
        # Querying today's records
        response = self.client.get(reverse('billing_records'))
        self.assertContains(response, "Today Customer")
        self.assertNotContains(response, "Past Customer")
        
        # Querying past date's records
        response = self.client.get(reverse('billing_records') + f"?date={past_date.strftime('%Y-%m-%d')}")
        self.assertContains(response, "Past Customer")
        self.assertNotContains(response, "Today Customer")

    def test_clear_due_action(self):
        """Verify POST to clear_due logs partial payments and clears balance when fully settled"""
        self.client.login(username="testuser", password="testpassword")
        
        # Create a bill with outstanding balance of $30.00
        bill = Bill.objects.create(
            customer_name="Due Customer", product=self.product,
            quantity=2, rate=45.00, total=90.00, amount_given=60.00, amount_to_be_given=-30.00
        )
        self.assertEqual(bill.amount_to_be_given, -30.00)
        
        # 1. Partial payment of $20.00
        response = self.client.post(reverse('clear_due', args=[bill.id]), {'clear_amount': '20.00'})
        bill_date_str = bill.created_at.strftime('%Y-%m-%d')
        self.assertRedirects(response, f"/billing/records/?date={bill_date_str}")
        
        bill.refresh_from_db()
        self.assertEqual(bill.amount_given, 80.00)
        self.assertEqual(bill.amount_to_be_given, -10.00)
        
        # 2. Final payment of $10.00 (settles balance)
        response = self.client.post(reverse('clear_due', args=[bill.id]), {'clear_amount': '10.00'})
        self.assertRedirects(response, f"/billing/records/?date={bill_date_str}")
        
        bill.refresh_from_db()
        self.assertEqual(bill.amount_given, 90.00)
        self.assertEqual(bill.amount_to_be_given, 0.00)
        self.assertEqual(bill.abs_amount_to_be_given, 0.00)

    def test_stock_pages_access_control(self):
        """Verify standard user is redirected from stock and stock1 views"""
        self.client.login(username="testuser", password="testpassword")
        
        # Access stock add -> should redirect to billing
        response = self.client.get(reverse('stock_add'))
        self.assertRedirects(response, reverse('billing'))
        
        # Access stock1 list -> should redirect to billing
        response = self.client.get(reverse('stock1'))
        self.assertRedirects(response, reverse('billing'))

    def test_stock_addition_and_listing(self):
        """Verify admin can add stock and it shows up in stock1 list"""
        self.client.login(username="testadmin", password="testpassword")
        
        # Access stock add page (this triggers category creation in forms.py)
        response = self.client.get(reverse('stock_add'))
        self.assertEqual(response.status_code, 200)
        
        category = Category.objects.get(name="Pesticide")
        
        # Submit stock addition
        response = self.client.post(reverse('stock_add'), {
            'category': category.id,
            'company_name': 'Test Metal Inc',
            'name': 'Stainless Steel Pipe',
            'description': '3-inch Grade 316',
            'stock_quantity': 25,
            'selling_price': 120.00,
            'stock_entered': '2026-06-09',
        })
        
        # Should redirect to stock_add
        self.assertRedirects(response, reverse('stock_add'))
        
        # Verify it shows up in stock_add page too
        response = self.client.get(reverse('stock_add'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Stainless Steel Pipe')
        
        # Verify Product was created with correct fields and defaults
        self.assertEqual(Product.objects.filter(name='Stainless Steel Pipe').count(), 1)
        product = Product.objects.get(name='Stainless Steel Pipe')
        self.assertEqual(product.company_name, 'Test Metal Inc')
        self.assertEqual(product.stock_quantity, 25)
        self.assertEqual(product.selling_price, 120.00)
        self.assertEqual(product.purchase_price, 120.00) # defaulted to selling price
        self.assertEqual(product.stock_entered, '2026-06-09')
        self.assertEqual(product.category.name, 'Pesticide') # submitted Pesticide category
        self.assertTrue(product.sku.startswith('STK-TEST-STAI-')) # auto-generated SKU
        
        # Verify it shows up in stock1 list page
        response = self.client.get(reverse('stock1'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Stainless Steel Pipe')
        self.assertContains(response, 'Test Metal Inc')
        self.assertContains(response, '2026-06-09')

    def test_admin_billing_counter_restricted(self):
        """Verify administrator is redirected when attempting to access billing counter views"""
        self.client.login(username="testadmin", password="testpassword")
        
        response = self.client.get(reverse('billing'))
        self.assertRedirects(response, reverse('dashboard'))
        
        response = self.client.get(reverse('billing_records'))
        self.assertRedirects(response, reverse('dashboard'))

    def test_stock_quantity_deduction_split(self):
        """Verify that counter billing reduces stock_quantity but keeps initial_quantity unchanged"""
        # 1. Log in as admin and create a product with qty = 50
        self.client.login(username="testadmin", password="testpassword")
        category = Category.objects.create(name="Seed")
        
        self.client.post(reverse('stock_add'), {
            'category': category.id,
            'company_name': 'Green Agri Co',
            'name': 'Hybrid Rice Seed',
            'description': 'High yielding hybrid rice seeds.',
            'stock_quantity': 50,
            'selling_price': 35.00,
            'stock_entered': '2026-06-09',
        })
        
        # Verify initial state on DB
        product = Product.objects.get(name='Hybrid Rice Seed')
        self.assertEqual(product.stock_quantity, 50)
        self.assertEqual(product.initial_quantity, 50)
        
        # 2. Log in as standard user to perform a billing transaction for 10 units
        self.client.login(username="testuser", password="testpassword")
        self.client.post(reverse('billing'), {
            'customer_name': 'Agri Client',
            'customer_address': 'Local Farm',
            'product': product.id,
            'quantity': 10,
            'rate': 35.00,
            'amount_given': 350.00
        })
        
        # Verify database fields
        product.refresh_from_db()
        self.assertEqual(product.stock_quantity, 40)       # Reduced: 50 - 10 = 40 (Stock 1)
        self.assertEqual(product.initial_quantity, 50)     # Unchanged: 50 (Stock)
