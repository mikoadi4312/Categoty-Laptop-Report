CREATE TABLE IF NOT EXISTS dim_brand (
    id SERIAL PRIMARY KEY,
    brand VARCHAR(100) UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS dim_store (
    id SERIAL PRIMARY KEY,
    id_store INT UNIQUE NOT NULL,
    store_name VARCHAR(255) NOT NULL
);

CREATE TABLE IF NOT EXISTS fact_sales (
    id SERIAL PRIMARY KEY,
    sale_date DATE NOT NULL,
    store_id INT REFERENCES dim_store(id),
    brand_id INT REFERENCES dim_brand(id),
    day_qty INT DEFAULT 0,
    day_revenue NUMERIC(18,2) DEFAULT 0,
    uploaded_at TIMESTAMP DEFAULT NOW(),
    uploaded_by VARCHAR(100),
    UNIQUE(sale_date, store_id, brand_id)
);

CREATE TABLE IF NOT EXISTS fact_stock (
    id SERIAL PRIMARY KEY,
    stock_date DATE NOT NULL,
    store_id INT REFERENCES dim_store(id),
    brand_id INT REFERENCES dim_brand(id),
    new_stock INT DEFAULT 0,
    demo_units INT DEFAULT 0,
    stock_volume INT DEFAULT 0,
    uploaded_at TIMESTAMP DEFAULT NOW(),
    uploaded_by VARCHAR(100),
    UNIQUE(stock_date, store_id, brand_id)
);

CREATE TABLE IF NOT EXISTS upload_log (
    id SERIAL PRIMARY KEY,
    upload_type VARCHAR(20),
    file_name VARCHAR(255),
    upload_date DATE,
    uploaded_by VARCHAR(100),
    rows_inserted INT,
    rows_updated INT,
    status VARCHAR(20),
    error_message TEXT,
    uploaded_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_sales_date ON fact_sales(sale_date);
CREATE INDEX IF NOT EXISTS idx_sales_store ON fact_sales(store_id);
CREATE INDEX IF NOT EXISTS idx_stock_date ON fact_stock(stock_date);
CREATE INDEX IF NOT EXISTS idx_stock_store ON fact_stock(store_id);
