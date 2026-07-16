from django import forms
from .models import Product, Category


class StockForm(forms.ModelForm):
    class Meta:
        model = Product
        fields = [
            'category', 'photo', 'company_name', 'name', 'size', 'description',
            'initial_quantity', 'selling_price', 'stock_entered',
            'vendor_cost', 'total_vendor_amount', 'amount_paid_to_vendor', 'remaining_amount_to_vendor', 'vendor_due_date'
        ]
        labels = {
            'category': 'Category',
            'photo': 'Photo (Optional)',
            'company_name': 'Company Name',
            'name': 'Product Name',
            'size': 'Unit / Size',
            'description': 'Description',
            'initial_quantity': 'Qty',
            'selling_price': 'Amt',
            'stock_entered': 'Stock Entered',
            'vendor_cost': 'Purchase Price (₹)',
            'total_vendor_amount': 'Total Vendor Amount (₹)',
            'amount_paid_to_vendor': 'Amount Paid to Vendor',
            'remaining_amount_to_vendor': 'Remaining Amount to Vendor',
            'vendor_due_date': 'Vendor Due Date',
        }
        widgets = {
            'stock_entered': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vendor_due_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Ensure allowed categories exist in database
        for cat_name in ["Pesticide", "Seed", "Fertilizers"]:
            Category.objects.get_or_create(name=cat_name)
            
        self.fields['category'].queryset = Category.objects.filter(name__in=["Pesticide", "Seed", "Fertilizers"])
        self.fields['category'].required = True
        self.fields['category'].empty_label = "-- Select Category --"

        # Enforce other required/optional states
        self.fields['company_name'].required = True
        self.fields['name'].required = True
        self.fields['description'].required = True
        self.fields['initial_quantity'].required = True
        self.fields['selling_price'].required = True
        self.fields['stock_entered'].required = True
        
        self.fields['photo'].required = False
        self.fields['size'].required = False
        self.fields['vendor_cost'].required = False
        self.fields['total_vendor_amount'].required = False
        self.fields['amount_paid_to_vendor'].required = False
        self.fields['remaining_amount_to_vendor'].required = False
        self.fields['vendor_due_date'].required = False

        # Override default 0 initial value for new products so field is blank
        if not self.instance.pk:
            self.fields['initial_quantity'].initial = None

