import sqlite3

def init_db(db_name="amazon_tracker.db"):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # 1. Ensure the table exists
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            asin TEXT PRIMARY KEY,
            region TEXT,
            brand TEXT,
            product_name TEXT,
            product_url TEXT,
            status TEXT,
            rating TEXT,
            review_count TEXT,
            coupon_info TEXT,
            ranking TEXT,
            listing_price TEXT,
            current_price TEXT,
            savings TEXT,
            screenshot_path TEXT,
            last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 2. Clear all previous records from the table so the dashboard starts fresh
    cursor.execute("DELETE FROM products")

    conn.commit()
    conn.close()
    print(f"Database '{db_name}' initialized and cleared successfully.")

def log_to_database(row, db_name="amazon_tracker.db"):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    cursor.execute('''
        INSERT INTO products (
            asin, region, brand, product_name, product_url,
            status, rating, review_count, coupon_info, ranking,
            listing_price, current_price, savings, screenshot_path, last_updated
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(asin) DO UPDATE SET
            region=excluded.region,
            brand=excluded.brand,
            product_name=excluded.product_name,
            product_url=excluded.product_url,
            status=excluded.status,
            rating=excluded.rating,
            review_count=excluded.review_count,
            coupon_info=excluded.coupon_info,
            ranking=excluded.ranking,
            listing_price=excluded.listing_price,
            current_price=excluded.current_price,
            savings=excluded.savings,
            screenshot_path=excluded.screenshot_path,
            last_updated=CURRENT_TIMESTAMP
    ''', (
        row.get("ID"), row.get("Region"), row.get("Brand", ""), row.get("Product Name", ""), row.get("Product URL"),
        row.get("Status"), row.get("Rating", ""), row.get("Review Count", ""),
        row.get("Coupon Info", ""), row.get("Ranking", ""), row.get("Listing Price", "N/A"), 
        row.get("Current Price", "N/A"), row.get("Savings", "0%"), row.get("Screenshot_Path", "")
    ))

    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()