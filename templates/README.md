# NovaMart — Full-Stack E-Commerce

A resume-ready full-stack project using **Python Flask** (backend) + **Excel/openpyxl** (database) + **Vanilla JS** (frontend).

---

## Tech Stack

| Layer       | Technology                  |
|-------------|-----------------------------|
| Backend     | Python 3 · Flask            |
| Database    | Excel (.xlsx) via openpyxl  |
| Frontend    | HTML5 · CSS3 · Vanilla JS   |
| Auth/Pay    | Mock (Luhn validation demo) |

---

## Project Structure

```
novamart/
├── app.py                  ← Flask REST API (backend)
├── create_db.py            ← Script to create Excel DB
├── data/
│   └── novamart_db.xlsx    ← Excel database (4 sheets)
├── templates/
│   ├── index.html          ← Product listing (home page)
│   ├── product.html        ← Product detail page
│   └── checkout.html       ← Cart + billing page
└── README.md
```

---

## Excel Database Sheets

| Sheet          | Purpose                              |
|----------------|--------------------------------------|
| `Products`     | Product catalog with stock levels    |
| `Orders`       | All placed orders with totals        |
| `PromoCodes`   | Discount codes (SAVE10, FREESHIP...) |
| `InventoryLog` | Stock change history per order       |

---

## API Endpoints

| Method | Endpoint           | Description              |
|--------|--------------------|--------------------------|
| GET    | `/api/products`    | List all products        |
| GET    | `/api/products/id` | Get one product          |
| POST   | `/api/promo`       | Validate promo code      |
| POST   | `/api/orders`      | Place order (writes Excel)|
| GET    | `/api/orders`      | List all orders (admin)  |

---

## Setup & Run

```bash
# 1. Install dependencies
pip install flask openpyxl

# 2. Create Excel database
python create_db.py

# 3. Start Flask server
python app.py

# 4. Open browser
# http://127.0.0.1:5000
```

---

## Promo Codes (demo)

| Code      | Effect                  |
|-----------|-------------------------|
| `SAVE10`  | 10% off subtotal        |
| `FREESHIP`| Free shipping           |
| `FLAT200` | ₹200 flat discount      |

---

## User Flow

1. **Home page** → Browse Amazon-style product cards with images, ratings, stock badges
2. **Click product image/name** → Product detail page (Add to Cart / Buy Now)
3. **Checkout** → Cart summary + payment form → Order saved to Excel

---

## Resume Points

- **RESTful API** with Flask — GET/POST endpoints, CORS, server-side validation
- **Excel as a database** — openpyxl for CRUD operations on `.xlsx` files
- **Multi-page SPA-style** frontend with localStorage cart persistence
- **Luhn algorithm** card validation
- **Inventory management** — stock deduction + audit log on every order
- **Promo code system** — percent/flat/free-shipping types, usage counter
- **Responsive design** — works on mobile and desktop
