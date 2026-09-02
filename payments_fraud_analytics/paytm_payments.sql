-- Step 1: Create Database Schema
CREATE DATABASE paytm_payments;
USE paytm_payments;

-- Step 2: Create Merchants Table
CREATE TABLE merchants (
    merchant_id INT PRIMARY KEY,
    merchant_name VARCHAR(100),
    category VARCHAR(50),
    region VARCHAR(50)
);

-- Step 3: Create Users Table
CREATE TABLE users (
    user_id INT PRIMARY KEY,
    signup_date DATETIME
);

-- Step 4: Create Transactions Table
CREATE TABLE transactions (
    transaction_id VARCHAR(50) PRIMARY KEY,
    user_id INT,
    merchant_id INT,
    transaction_time DATETIME,
    amount_inr DECIMAL(10,2),
    payment_method VARCHAR(50),
    status VARCHAR(50),
    risk_score INT,
    FOREIGN KEY (user_id) REFERENCES users(user_id),
    FOREIGN KEY (merchant_id) REFERENCES merchants(merchant_id)
);

-- Verification tables
SELECT 'merchants' AS table_name, COUNT(*) AS total_rows FROM merchants
UNION ALL
SELECT 'users', COUNT(*) FROM users
UNION ALL
SELECT 'transactions', COUNT(*) FROM transactions;

-- Merchant Regional Volume Breakdown
SELECT 
    m.region,
    COUNT(t.transaction_id) AS total_txns,
    SUM(CASE WHEN t.status = 'captured' THEN t.amount_inr ELSE 0 END) AS captured_gmv_inr
FROM merchants m
LEFT JOIN transactions t ON m.merchant_id = t.merchant_id
GROUP BY m.region
ORDER BY captured_gmv_inr DESC;

-- Payment Method Success Rates Calculation
SELECT 
    payment_method,
    COUNT(DISTINCT transaction_id) AS total_transactions,
    COUNT(DISTINCT user_id) AS unique_users,
    SUM(CASE WHEN status = 'captured' THEN 1 ELSE 0 END) AS successful_txns,
    ROUND(SUM(CASE WHEN status = 'captured' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2) AS success_rate_pct
FROM transactions
WHERE status IN ('captured', 'failed')
GROUP BY payment_method
ORDER BY total_transactions DESC;

-- Top 5 Highest Risk Score Transactions
SELECT 
    t.transaction_id,
    m.merchant_name,
    t.amount_inr,
    t.payment_method,
    t.risk_score,
    t.status
FROM transactions t
INNER JOIN merchants m ON t.merchant_id = m.merchant_id
WHERE t.risk_score >= 90
ORDER BY t.risk_score DESC, t.amount_inr DESC
LIMIT 5;

-- High-Risk Transaction Breakdown by Merchant Category
SELECT 
    m.category,
    COUNT(t.transaction_id) AS high_risk_txn_count,
    ROUND(AVG(t.risk_score), 2) AS avg_risk_score,
    SUM(t.amount_inr) AS high_risk_exposure_inr
FROM transactions t
INNER JOIN merchants m ON t.merchant_id = m.merchant_id
WHERE t.risk_score >= 70
GROUP BY m.category
ORDER BY high_risk_txn_count DESC, high_risk_exposure_inr DESC;

-- Quantifying Chargeback Impact
SELECT 
    COUNT(transaction_id) AS total_chargeback_txns,
    COUNT(DISTINCT user_id) AS unique_users_affected,
    SUM(amount_inr) AS total_chargeback_amount_inr
FROM transactions
WHERE status = 'chargeback';

-- Detecting Burner Accounts 
SELECT 
    t.transaction_id,
    t.user_id,
    u.signup_date,
    t.transaction_time,
    DATEDIFF(t.transaction_time, u.signup_date) AS account_age_days,
    t.amount_inr,
    t.status
FROM transactions t
INNER JOIN users u ON t.user_id = u.user_id
WHERE t.status = 'chargeback'
  AND DATEDIFF(t.transaction_time, u.signup_date) BETWEEN 0 AND 29
ORDER BY t.transaction_time ASC;

-- Detecting Velocity Attacks
SELECT 
    user_id,
    FROM_UNIXTIME(FLOOR(UNIX_TIMESTAMP(transaction_time) / 600) * 600) AS time_bucket_10m,
    COUNT(transaction_id) AS transaction_count,
    SUM(amount_inr) AS total_amount_inr
FROM transactions
GROUP BY user_id, time_bucket_10m
HAVING COUNT(transaction_id) >= 3
ORDER BY time_bucket_10m ASC;

