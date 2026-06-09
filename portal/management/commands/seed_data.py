from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from django.utils import timezone
import datetime
import random

from portal.models import Category, Product, Customer, Order, OrderItem

class Command(BaseCommand):
    help = "Seeds database with initial category, product, customer, and order records for demonstration."

    def handle(self, *args, **options):
        self.stdout.write("Starting data seeding process...")

        # 1. Create Superuser if not exists
        if not User.objects.filter(username="admin").exists():
            User.objects.create_superuser("admin", "admin@rukminienterprises.com", "admin123")
            self.stdout.write(self.style.SUCCESS("Created superuser 'admin' with password 'admin123'."))
        else:
            self.stdout.write("Superuser 'admin' already exists.")

        # Create Standard User if not exists
        if not User.objects.filter(username="user").exists():
            User.objects.create_user("user", "user@rukminienterprises.com", "user123")
            self.stdout.write(self.style.SUCCESS("Created standard user 'user' with password 'user123'."))
        else:
            self.stdout.write("Standard user 'user' already exists.")

        # Clear existing data to prevent duplicate primary keys or SKU collision
        OrderItem.objects.all().delete()
        Order.objects.all().delete()
        Product.objects.all().delete()
        Customer.objects.all().delete()
        Category.objects.all().delete()

        # 2. Create Categories
        cats_data = [
            {"name": "Electrical Supplies", "description": "Cables, reels, fittings, switches, and wiring accessories."},
            {"name": "Metal Products", "description": "Steel plates, rods, channels, brass, and aluminum profiles."},
            {"name": "Power Tools & Machinery", "description": "Industrial drills, cutters, lathes, and heavy equipment."},
            {"name": "Fasteners & Hardware", "description": "Screws, bolts, anchors, hinges, and structural brackets."}
        ]
        
        categories = {}
        for cat in cats_data:
            c = Category.objects.create(name=cat["name"], description=cat["description"])
            categories[cat["name"]] = c
            self.stdout.write(f"Created Category: {c.name}")

        # 3. Create Products
        products_data = [
            {
                "name": "Copper Wire Reel 100m", "sku": "COP-WIR-100", 
                "category": categories["Electrical Supplies"], "purchase_price": 45.00, 
                "selling_price": 75.00, "stock_quantity": 25, "min_stock_level": 10,
                "description": "High-conductivity copper cable, 2.5mm thickness, flexible insulation."
            },
            {
                "name": "Steel Sheet 2mx1m (3mm)", "sku": "STL-SHT-3MM", 
                "category": categories["Metal Products"], "purchase_price": 85.00, 
                "selling_price": 149.00, "stock_quantity": 4, "min_stock_level": 8,
                "description": "Hot-rolled mild steel structural sheet, perfect for fabrication."
            },
            {
                "name": "Heavy-Duty Bench Drill Press", "sku": "HVY-DRL-PRS", 
                "category": categories["Power Tools & Machinery"], "purchase_price": 380.00, 
                "selling_price": 699.00, "stock_quantity": 2, "min_stock_level": 3,
                "description": "16-speed floor-standing drill press with 550W induction motor."
            },
            {
                "name": "Brass Wood Screws (Box of 500)", "sku": "BRS-SCR-500", 
                "category": categories["Fasteners & Hardware"], "purchase_price": 7.20, 
                "selling_price": 16.50, "stock_quantity": 65, "min_stock_level": 15,
                "description": "Flat-head countersunk brass screws, 4.0mm x 30mm."
            },
            {
                "name": "Aluminum Round Tubing 3m", "sku": "ALM-TUB-3M", 
                "category": categories["Metal Products"], "purchase_price": 18.50, 
                "selling_price": 38.00, "stock_quantity": 5, "min_stock_level": 10,
                "description": "6063-T6 aluminum extrusion tube, 25mm outer diameter, 2mm wall."
            },
            {
                "name": "Industrial Angle Grinder 9-inch", "sku": "IND-AGR-9IN", 
                "category": categories["Power Tools & Machinery"], "purchase_price": 95.00, 
                "selling_price": 179.00, "stock_quantity": 12, "min_stock_level": 5,
                "description": "High-torque 2200W motor with safety clutch and soft start."
            },
            {
                "name": "Zinc Plated Bolts M10 (Pack of 100)", "sku": "ZNC-BLT-M10", 
                "category": categories["Fasteners & Hardware"], "purchase_price": 12.00, 
                "selling_price": 28.00, "stock_quantity": 30, "min_stock_level": 8,
                "description": "Hex head grade 8.8 medium carbon steel bolts, zinc-plated."
            }
        ]

        products = []
        for prod in products_data:
            p = Product.objects.create(**prod)
            products.append(p)
            self.stdout.write(f"Created Product: {p.name}")

        # 4. Create Customers
        customers_data = [
            {"name": "Rohan Sharma", "company_name": "Sharma Metal Works", "email": "rohan@sharmametals.in", "phone": "+91 98765 43210", "address": "Plot 42, Sector 5, Industrial Area, Noida"},
            {"name": "Priya Verma", "company_name": "Verma Construction Group", "email": "pverma@vermacon.com", "phone": "+91 99123 45678", "address": "Verma Towers, Connaught Place, New Delhi"},
            {"name": "Tarun Gupta", "company_name": "Gupta Hardware & Electricals", "email": "sales@guptahardware.com", "phone": "+91 98112 23344", "address": "Shop 12, Main Bazar, Gurgaon"},
            {"name": "Anil Mehta", "company_name": "", "email": "anilmehta@gmail.com", "phone": "+91 95600 11223", "address": "H.No 124, Sector 15, Faridabad"}
        ]

        customers = []
        for cust in customers_data:
            c = Customer.objects.create(**cust)
            customers.append(c)
            self.stdout.write(f"Created Customer: {c.name}")

        # 5. Create Historical Orders
        today = timezone.now()
        
        # We will create orders representing the last 6 months (including the current month)
        # We save the orders first, then update their order_date directly with ORM update to bypass auto_now_add.
        self.stdout.write("Seeding historical invoice records...")
        
        for month_offset in range(5, -1, -1):
            # Calculate target date for the order
            target_date = today - datetime.timedelta(days=month_offset * 30 + random.randint(1, 10))
            
            # Select random customer and status
            customer = random.choice(customers)
            # Most past orders are Paid or Shipped; current month might have Pending
            status = "Paid" if month_offset > 0 else random.choice(["Paid", "Pending", "Shipped"])
            
            # Create Order
            order = Order.objects.create(
                customer=customer,
                status=status,
                notes=f"Demonstration order generated for {target_date.strftime('%B %Y')}"
            )
            
            # Select 1-2 random products to add to this order
            sampled_products = random.sample(products, random.randint(1, 2))
            for prod in sampled_products:
                qty = random.randint(1, 4)
                OrderItem.objects.create(
                    order=order,
                    product=prod,
                    quantity=qty,
                    price=prod.selling_price
                )
            
            # Bypass auto_now_add using ORM update
            Order.objects.filter(pk=order.pk).update(order_date=target_date)
            
        self.stdout.write(self.style.SUCCESS("Database seeded successfully with all mock elements!"))
