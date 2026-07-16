# Rukmini Enterprises ERP - Billing & Inventory Management System

Welcome to the **Rukmini Enterprises ERP** system. This is a comprehensive, web-based management portal designed to streamline billing, stock inventory, vendor payments, and expenses. Built on Python and the Django web framework, it offers advanced reporting, PDF/Excel exporting, FIFO batch tracking, and access-control features.

---

## 🚀 Key Features

### 1. User Dashboard & Role-based Access
- **Admin Dashboard**: Full visibility into overall sales, pending customer dues, recent vendor payments, overall inventory metrics, and expense tracking.
- **Standard User Dashboard**: Restrictive dashboard allowing staff to focus on checkout, billing, and day-to-day operations.
- **Custom Middleware**: Extra security layers (e.g., GST restriction controls) to restrict certain features based on user permission policies.

### 2. Billing & Invoicing
- **Interactive Checkout**: Create bills/invoices with dynamic product selection, quantity tracking, and instant totals calculation.
- **Dues Management**: Support for partial payments. The system computes outstanding customer dues and lets administrators easily clear/update dues later.
- **Automated Calculations**: Dynamic 18% GST computation and option to toggling GST pricing logic.
- **Invoice Export**: Download single bills directly as clean, formatted PDF invoices.

### 3. Inventory & Stock Management
- **FIFO Batch Tracking**: Products can be tracked via discrete batches (`ProductBatch`), supporting batch-specific purchase rates, selling prices, and FIFO (First In, First Out) inventory depletion on sales.
- **Low Stock Alerts**: Automatically highlights products falling below their minimum stock thresholds.
- **Stock Log History**: Logs additions to inventory per product (`StockHistory`) for transparency.
- **Stock Ledger**: Review complete audit trails of stock changes.

### 4. Vendor & Expenses Management
- **Vendor Ledger**: Track details of products supplied by specific vendors, total vendor cost, amount paid, remaining balance, and due dates.
- **Vendor Payments**: Log and track vendor payment history (`VendorPaymentHistory`) to ensure clear accounting records.
- **General Expenses**: Track company-wide expenses (utility, logistics, rent) paid out, associated with logged users.

### 5. Advanced Reports & Exports
- **Excel Downloads**: Export dashboards, revenue summaries, product catalogs, and vendor ledgers to spreadsheet formats (via `openpyxl`).
- **PDF Reports**: Generate structured reports for printing and sharing (via `reportlab`).

---

## 🛠️ Tech Stack

- **Backend**: Python 3, Django Web Framework (>= 4.2)
- **Database**: SQLite (default configuration)
- **PDF Generation**: ReportLab
- **Excel Spreadsheet Handling**: OpenPyXL
- **Frontend**: Responsive HTML5, Vanilla CSS, Django Template Engine

---

## 📋 Installation & Setup

Follow these steps to run the application locally:

### 1. Clone the Repository
```bash
git clone https://github.com/tarun4567/Rukmini-Enterprises.git
cd Rukmini-Enterprises
```

### 2. Set Up a Virtual Environment
Create and activate a virtual environment for the project:
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS/Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Run Database Migrations
Apply the existing database schema to your local database:
```bash
python manage.py migrate
```

### 5. Create a Superuser (Admin Account)
To log in and access administrative features, create an admin account:
```bash
python manage.py createsuperuser
```
Follow the prompts to configure your username, email, and password.

### 6. Run the Development Server
```bash
python manage.py runserver
```
Open your browser and navigate to `http://127.0.0.1:8000/` to access the application. Go to `http://127.0.0.1:8000/admin/` to manage raw database entries.

---

## 📁 Project Directory Structure

```text
├── manage.py                  # Django CLI entrypoint
├── requirements.txt           # Python dependencies
├── db.sqlite3                 # Local database (SQLite)
├── rukmini_erp/               # Core Django project settings & main URL configuration
│   ├── settings.py
│   ├── urls.py
│   └── wsgi.py
└── portal/                    # Main application directory
    ├── models.py              # Database models (Product, Bill, Expense, batches, etc.)
    ├── views.py               # View controllers, PDF, and Excel exports
    ├── forms.py               # Custom Django forms for stock, expenses, and checkout
    ├── urls.py                # App-specific path routing
    ├── middleware.py          # Custom middleware logic (e.g., GST restrict rules)
    ├── templates/             # HTML Templates (Billing, stock, vendor sheets, dashboards)
    └── migrations/            # Database schema migration files
```

---

## 📜 License & Contributions
This project is open-source. For contributions, bug reports, or feature requests, please open an issue or submit a pull request to the repository.
