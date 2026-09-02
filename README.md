**----------------------------------------------------Paytm FinTech Analytics & AI Platform----------------------------------------------**
**===================================================================================================**
                                          **Part 1 — Payments & Fraud Analytics (/payments_fraud_analytics)**
**===================================================================================================**
    
    **Part A — Excel/Sheets merchant workbook**
        **Module Requirements & Scope**
            Part A requires the creation of a structured Excel financial workbook (merchant_workbook.xlsx) alongside sepcific specifications. The workbook must fulfill four core data processing and analytics criteria:
            1. Merge merchant data (merchant_name, category, region) from merchants.csv into a transaction-level ledger sheet using fixed absolute range references ($) and wrap them in IFERROR/IFNA to output "Merchant not found" for unmatched IDs.
            2. Construct a horizontal reference table defining payment method MDR (Merchant Discount Rate) fee percentages and lookup each transaction’s fee rate dynamically.
            3. Build a conditional classification rule labeling transactions as "High-Value Merchant Day" based on merchant daily volume and regional filtering.
            4. Summarize transaction amounts and counts by merchant_id and status, comparing total transaction counts against distinct transacted calendar days for at least 5 merchants.
        
        **Detailed Workbook Architecture (merchant_workbook.xlsx)**
            The workbook consists of five interconnected sheets:
            1. Ledger: Raw transaction data imported from ledger.csv.
            2. Merchants: Merchant data imported from merchants.csv.
            3. MDR Rates: Horizontally laid out lookup matrix for payment gateway fees.
            4. Transactions-view sheet: Primary consolidated analytical ledger featuring formulas for lookups, daily aggregations, and high-value risk flags.
            5. Pivot Table: Multidimensional analytical summary layer providing daily volume rollups and merchant activity metrics.

        **Implementation and Analysis**
            1. Fixed Range VLOOKUP Implementation:
                To pull data from the Merchants sheet into the Transactions-view sheet, absolute column and row references ($A$2:$D$41) were applied. Wrapping the formula with IFERROR ensures robust error handling if an invalid merchant_id is encountered. 
                    Merchant Name Lookup (Column I): =IFERROR(VLOOKUP(C2,Merchants!$A$2:$D$41,2,FALSE),"Merchant not found")
                    Category Lookup (Column J): =IFERROR(VLOOKUP(C2, Merchants!$A$2:$D$41, 3, FALSE), "Merchant not found")
                    Region Lookup (Column K): =IFERROR(VLOOKUP(C2, Merchants!$A$2:$D$41, 4, FALSE), "Merchant not found")
                
                Explanation: C2 represents the transaction's merchant_id. The absolute range Merchants!$A$2:$D$41 locks the reference table across all 547 transaction rows, preventing lookup degradation during drag-down copying.

            2. HLOOKUP MDR Fee Tier Matrix:
                The MDR Rates sheet hosts a horizontal reference grid mapping payment channels to MDR percentage tiers:
                    | Reference Header | Col 1 | Col 2  | Col 3 | Col 4      |
                    | Payment Method   | UPI   | Wallet | Card  | Netbanking |
                    | MDR Fee %        | 0.00% | 1.00%  | 2.00% | 1.50%      |
                In Column L of the Transactions-view sheet, HLOOKUP evaluates the payment method (F2) against the horizontally oriented lookup range MDR Rates!$C$1:$F$2 with the following formula:
                     =IFNA(HLOOKUP(F2, 'MDR Rates'!$C$1:$F$2, 2, FALSE), "N/A")

                Explanation: If a transaction uses "Card", the formula scans Row 1 horizontally to locate "Card" in Column E and returns 2.00% from Row 2.

            3. Nested IF/AND Classification Logic:
                To detect high-velocity merchant days, daily aggregate merchant transaction volumes were derived using two methods:
                    1. SUMIFS Aggregation (Column N): =SUMIFS($E$2:$E$548, $C$2:$C$548, C2, $M$2:$M$548, M2)
                    2. GETPIVOTDATA Aggregation (Column O): =GETPIVOTDATA("amount_inr", 'Pivot Table'!$A$3, "merchant_name", I2, "Date", M2)

                Documented Classification Rule: A transaction is flagged as "High-Value Merchant Day" if:
                    1. The merchant's total aggregate volume for that specific calendar date exceeds INR 5,000 (O2 > 5000).
                    2. AND the merchant's operating region is NOT "East" (K2 <> "East").

                Formula: =IF(AND(O2 > 5000, K2 <> "East"), "High-Value Merchant Day", "Standard")

            4. Pivot Table Summary & Activity Ratios:
                The Pivot Table sheet summarizes payment processing metrics across status categories (captured, chargeback, failed) and analyzes merchant trading frequency.

                    **Pivot Table 1: Daily Merchant Aggregation (Left Side):**
                    Purpose - Aggregates daily transaction volume per merchant to feed the GETPIVOTDATA formula used for high-value day flagging.

                        Rows: Date and merchant_name
                        Values: amount_inr and set its summarization to Sum.

                    Analysis: The pivot table evaluates cash flow velocity by collapsing transaction streams into daily merchant totals across the month of Jan 2026. Across the portfolio, daily merchant processing volume averages ₹863.66, with a median of ₹299.00, indicating that transaction flows are predominantly driven by steady, low-ticket retail activities. Applying the business logic cut-off—flagging days where daily merchant turnover exceeds ₹5,000 and operating region excludes "East"—isolates 4 outlier transaction clusters (Merchant_001, Merchant_017, Merchant_022 and Merchant_039). These high-value surges reach up to ₹9,796 on single calendar dates (such as Merchant_039 in the North region on January 22, 2026). 

                    **Pivot Table 2: Merchant Status & Activity Ratio Summary (Right Side):**
                    Purpose - Summarizes financial volume, overall transaction counts, and unique transacted days per merchant split by transaction status.

                        Rows & Columns:
                            * Rows: merchant_id
                            * Columns: status (e.g., captured, chargeback, failed)

                        Values (Count vs. Distinct Count):
                            * Date summarizing values by Distinct Count to count unique active transacted days.
                            * transaction_id summarized as Count to measure total transaction volume.
                            * amount_inr summarized as Sum to obtain total monetary volume.

                    Analysis: The second pivot table evaluates portfolio health and trading consistency across all 40 active merchants by segmenting total transacted amounts, total transaction counts, and distinct active calendar days by payment status. Across the overall dataset, transaction statuses are dominated by successful settlements, with ₹290,382 captured across 468 transactions, compared to ₹54,472 in chargebacks (28 transactions) and ₹37,749 in failed processing (51 transactions). Merchants exhibiting consistent daily activity across multiple calendar days demonstrate organic retail throughput, whereas entities exhibiting low distinct transacted days paired with high chargeback or failure counts (such as merchants accumulating multiple failed processing attempts within 1 to 2 days) indicate authorization friction, integration failures, or potential chargeback abuse that requires intervention by platform risk teams.
    
    **Part B — SQL fraud-pattern detection:**
        **Module Requirements & Scope**
            The objective of Part B is to construct a relational schema in SQLite (paytm_payments.db), ingest raw payment datasets (merchants.csv, users.csv, ledger.csv), and execute targeted SQL queries to expose systemic fraud vectors across the transaction network. 

            Functional Specifications:
                1. Schema: Model entities with strict Primary Key (PK) and Foreign Key (FK) constraints: merchants (PK merchant_id), users (PK user_id), and transactions (PK transaction_id, FK user_id, FK merchant_id).	
                2. SQL Operator Coverage: Deliver queries utilizing SELECT, WHERE, ORDER BY, LIMIT, DISTINCT, GROUP BY, HAVING, INNER JOIN, and LEFT JOIN.	
                3. Chargeback Impact Quantification: Compute total chargeback volume, unique affected users, and total monetary exposure resulting from disputed transactions.	
                4. Burner Account Detection: Surface accounts created within 30 days prior to issuing a chargeback transaction, obeying the strict boundary (transaction_time - signup_date) < 30 days (surfacing all 15 seeded burner account rows).	
                5. Velocity Attack Detection: Detect short-window transaction floods by identifying users initiating transactions within any 10-minute sliding/floored time bucket (surfacing all 8 seeded velocity clusters).

        **Detailed Architecture Design (paytm_payments.sql)**
        Table Definitions & Primary Key / Foreign Key Mappings
            * merchants: Stores partner merchant master attributes.
                - Primary Key: merchant_id
            * users: Stores user registration profiles.
                - Primary Key: user_id
            * transactions: Transactions table recording individual payment attempts.
                - Primary Key: transaction_id
                - Foreign Keys: user_id refers users(user_id), merchant_id refers merchants(merchant_id)
       
        **Implementation and Analysis**
        
        Step 1: Database Setup & Data Ingestion
        The relational schema is created, and CSV datasets are populated into paytm_payments.db.
            CREATE TABLE merchants (
                merchant_id INT PRIMARY KEY,
                merchant_name VARCHAR(100),
                category VARCHAR(50),
                region VARCHAR(50)
            );

            CREATE TABLE users (
                user_id INT PRIMARY KEY,
                signup_date DATETIME
            );

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

        Step 2: Key SQL Queries & Executed Outputs
        Query 1: Database Row Verification (UNION ALL)
            SELECT 'merchants' AS table_name, COUNT(*) AS total_rows FROM merchants
            UNION ALL
            SELECT 'users', COUNT(*) FROM users
            UNION ALL
            SELECT 'transactions', COUNT(*) FROM transactions;
        
        Output:
            | table_name    | total_rows |
            | merchants     | 40         |
            | users         | 365        |
            | transactions  | 547        |

        Analysis: Executing a UNION ALL query across the three core tables provides structural validation and confirms data integrity following database ingestion. The schema verification confirms an ingestion total of 40 merchants, 365 unique registered users, and 547 logged transaction records. By confirming that all primary entities are fully populated without record truncation, this check ensures that downstream relational joins, foreign key enforcement, and aggregate fraud detection queries operate on a complete and uncorrupted dataset.

        Query 2: Regional Merchant Volume Breakdown (LEFT JOIN, GROUP BY)
            SELECT 
                m.region,
                COUNT(t.transaction_id) AS total_txns,
                SUM(CASE WHEN t.status = 'captured' THEN t.amount_inr ELSE 0 END) AS captured_gmv_inr
            FROM merchants m
            LEFT JOIN transactions t ON m.merchant_id = t.merchant_id
            GROUP BY m.region
            ORDER BY captured_gmv_inr DESC;
        
        Output:
            | region | total_txns | captured_gmv_inr |
            | East   |        180 |       100,096.00 |
            | North  |        154 |        77,773.00 |
            | South  |        127 |        73,989.00 |
            | West   |         86 |       ₹38,524.00 |

        Analysis: Analyzing Gross Merchandise Value (GMV) by geographic operating region reveals significant concentration of successful transaction volume in the East region, which leads the portfolio with ₹100,096 across 180 total transactions. North and South regions display highly comparable performance levels, generating ₹77,773 (154 transactions) and ₹73,989 (127 transactions) in captured GMV respectively. In contrast, the West region exhibits lower adoption and processing velocity, accounting for ₹38,524 across 86 transactions. This breakdown demonstrates that revenue distribution is heavily weighted toward Eastern commercial hubs, while the West represents an underpenetrated market requiring targeted merchant acquisition strategies.

        Query 3: Payment Method Success Rates (WHERE, DISTINCT, ROUND)
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

        Output:
            | payment_method | total_transactions | unique_users | successful_txns | success_rate_pct |
            | UPI            |                266 |          186 |             255 |           95.86% |
            | Card           |                105 |           70 |              76 |           72.38% |
            | Wallet         |                 95 |           85 |              87 |           91.58% |
            | Netbanking    |                  53 |           50 |              50 |           94.34% |

        Analysis: Evaluating payment channel performance across transactions highlights Unified Payments Interface (UPI) as both the dominant payment rails and the most reliable processing channel. UPI processed 266 total transactions across 186 unique users with an industry-leading success rate of 95.86% (255 captured transactions). Netbanking and Wallet payment methods also demonstrated strong technical stability, yielding success rates of 94.34% (50 of 53 transactions) and 91.58% (87 of 95 transactions) respectively. Conversely, Card transactions experienced high authorization friction, registering a noticeably lower success rate of 72.38% (76 captured out of 105 attempts), pointing to potential issuer bank drop-offs, authentication failures, or stricter gateway fraud rules applied to card instruments.


        Query 4: Top 5 Highest Risk Score Transactions (INNER JOIN, ORDER BY, LIMIT)
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

        Output:
            | transaction_id | merchant_name | amount_inr | payment_method | risk_score | status     |
            |   TXN200004    |  Merchant_019 |   4,999.00 | Card           |        100 | chargeback |
            |   TXN100156    |  Merchant_020 |   2,999.00 | UPI            |        100 | captured   |
            |   TXN200007    |  Merchant_036 |     999.00 | Card           |        100 | chargeback |
            |   TXN100436    |  Merchant_025 |     799.00 | UPI            |        100 | captured   |
            |   TXN100498    |  Merchant_031 |     499.00 | Wallet         |        100 | captured   |

        Analysis: Filtering for critical risk scores highlights transactions flagged at the maximum security threshold of 100. The top high-risk events are characterized by binary outcomes: while high-value card payments such as TXN200004 (₹4,999 on Merchant_019) ultimately resulted in chargebacks, several other maximum-risk transactions (such as TXN100156 for ₹2,999 via UPI) were successfully captured. This operational pattern indicates that elevated risk scores independently correlate with disputed outcomes on credit card rails, whereas high-risk UPI transactions often pass automated risk scoring without immediate failure, exposing potential blind spots in real-time fraud scoring models.

        Query 5: High-Risk Transaction Breakdown by Merchant Category (COUNT, AVG, SUM, INNER JOIN, GROUP BY, ORDER BY)
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

        Output:
            | category      | high_risk_txn_count | avg_risk_score  | high_risk_exposure_inr   |
            | grocery       |           44        |       85.07     |          34606.00        |
            | travel        |           38        |       84.24     |          32712.00        |
            | ecommerce     |           38        |       83.95     |          27962.00        |
            | food_delivery |           33        |       84.45     |          23567.00        |
            | entertainment |           23        |       83.43     |          21277.00        |
            | bill_payment  |           14        |       84.21     |          10986.00        |
            | recharge      |            9        |       77.22     |          11091.00        |

        Analysis: Across the evaluated spending categories, Grocery and Travel represent the highest concentration of financial risk, generating ₹34,606.00 (44 transactions) and ₹32,712.00 (38 transactions) in high-risk exposure respectively, while maintaining peak severity scores above 84.0. Across most categories (including E-commerce, Food Delivery, and Bill Payments), the average risk score remains tightly clustered between 83.43 and 85.07, indicating that flagged transactions across daily merchant touchpoints share consistently high threat profiles. Conversely, Recharge exhibits the lowest overall risk profile with an average risk score of 77.22 across only 9 flagged transactions. To protect platform liquidity effectively, real-time defenses should prioritize velocity monitoring on high-volume everyday spend while enforcing step-up authentication on high-risk bill payments before settlement.

        Query 6: Quantifying Chargeback Impact (COUNT, DISTINCT, SUM)
            SELECT 
                COUNT(transaction_id) AS total_chargeback_txns,
                COUNT(DISTINCT user_id) AS unique_users_affected,
                SUM(amount_inr) AS total_chargeback_amount_inr
            FROM transactions
            WHERE status = 'chargeback';

        Output:
            | total_chargeback_txns | unique_users_affected | total_chargeback_amount_inr |
            |          28           |           27          |           54,472.00         |
       
        Analysis: Quantifying overall chargeback exposure confirms a total loss of ₹54,472 spread across 28 disputed transactions and affecting 27 unique users. The close alignment between transaction count (28) and unique user count (27) indicates that chargebacks are not concentrated among a handful of multi-dispute accounts, but are instead broadly distributed across distinct user entities acting independently. 

        Query 7: Burner Account Detection (INNER JOIN, DATEDIFF / Day-Difference Boundary)
            SELECT 
                t.transaction_id,
                t.user_id,
                u.signup_date,
                t.transaction_time,
                CAST((julianday(t.transaction_time) - julianday(u.signup_date)) AS INT) AS account_age_days,
                t.amount_inr,
                t.status
            FROM transactions t
            INNER JOIN users u ON t.user_id = u.user_id
            WHERE t.status = 'chargeback'
            AND (julianday(t.transaction_time) - julianday(u.signup_date)) >= 0
            AND (julianday(t.transaction_time) - julianday(u.signup_date)) < 30
            ORDER BY t.transaction_time ASC;

        Output:
            | transaction_id | user_id | signup_date | transaction_time | account_age_days | amount_inr | status     |
            |  TXN200001     |     352 | 2025-12-31 12:00:00 | 2026-01-11 12:00:00 |   11  |  4,999.00  | chargeback |
            |  TXN200009     |     360 | 2025-12-22 13:00:00 | 2026-01-13 13:00:00 |   22  |  1,999.00  | chargeback |
            |  TXN200004     |     355 | 2026-01-05 12:00:00 | 2026-01-16 12:00:00 |   11  |  4,999.00  | chargeback |
            |  TXN200014     |     365 | 2025-12-27 21:00:00 | 2026-01-18 21:00:00 |   22  |  1,999.00  | chargeback |
            |  TXN200010     |     361 | 2026-01-11 07:00:00 | 2026-01-20 07:00:00 |    9  |  4,999.00  | chargeback |
            |  TXN200002     |     353 | 2026-01-10 14:00:00 | 2026-01-21 14:00:00 |   11  |  1,999.00  | chargeback |
            |  TXN200003     |     354 | 2025-12-29 19:00:00 | 2026-01-21 19:00:00 |   23  |  4,999.00  | chargeback |
            |  TXN200013     |     364 | 2026-01-04 22:00:00 | 2026-01-22 22:00:00 |   18  |    999.00  | chargeback |
            |  TXN200011     |     362 | 2026-01-08 02:00:00 | 2026-01-23 02:00:00 |   15  |  4,999.00  | chargeback |
            |  TXN200006     |     357 | 2026-01-19 11:00:00 | 2026-01-23 11:00:00 |    4  |  1,999.00  | chargeback |
            |  TXN200012     |     363 | 2026-01-06 17:00:00 | 2026-01-23 17:00:00 |   17  |    999.00  | chargeback |
            |  TXN200008     |     359 | 2026-01-18 22:00:00 | 2026-01-25 22:00:00 |    7  |  2,999.00  | chargeback |
            |  TXN200007     |     358 | 2026-01-06 05:00:00 | 2026-01-28 05:00:00 |   22  |    999.00  | chargeback |
            |  TXN200005     |     356 | 2026-01-18 07:00:00 | 2026-01-29 07:00:00 |   11  |  2,999.00  | chargeback |
            |  TXN200000     |     351 | 2026-01-15 06:00:00 | 2026-01-30 06:00:00 |   15  |  1,999.00  | chargeback |

        Analysis: Isolating chargeback transactions tied to newly established profiles (account age strictly less than 30 days) surfaces all 15 seeded burner account incidents, exposing a targeted onboarding exploitation vector. These 15 burner accounts represent 53.57% of total portfolio chargeback incidents and account for ₹43,985 (80.75%) of total monetary chargeback losses. With an average account lifespan of only 14.5 days prior to dispute initiation and an elevated average transaction amount of ₹2,932.33, the data confirms that bad actors systematically register accounts, execute single high-value purchases, and file payment disputes before identity verification procedures mature.

        Query 8: Velocity Attack Detection (10-Minute Time Bucketing, HAVING >= 3)
            SELECT 
                user_id,
                datetime((strftime('%s', transaction_time) / 600) * 600, 'unixepoch') AS time_bucket_10m,
                COUNT(transaction_id) AS transaction_count,
                SUM(amount_inr) AS total_amount_inr
            FROM transactions
            GROUP BY user_id, time_bucket_10m
            HAVING COUNT(transaction_id) >= 3
            ORDER BY time_bucket_10m ASC;

        Output:
            | user_id |   time_bucket_10m   | transaction_count | total_amount_inr |
            | 200     | 2026-01-01 22:00:00 |                 4 |         1,396.00 |
            | 314     | 2026-01-02 18:00:00 |                 4 |         1,496.00 |
            | 154     | 2026-01-02 22:00:00 |                 4 |         1,596.00 |
            | 59      | 2026-01-09 21:00:00 |                 4 |         1,596.00 |
            | 73      | 2026-01-12 09:00:00 |                 4 |         1,496.00 |
            | 229     | 2026-01-12 12:00:00 |                 4 |         1,496.00 |
            | 287     | 2026-01-14 14:00:00 |                 4 |         1,696.00 |
            | 345     | 2026-01-23 09:00:00 |                 4 |         1,696.00 |

        Analysis: Applying a 10-minute floored time window isolates all 8 seeded velocity attack clusters, identifying rapid-fire transaction floods executed by users 59, 73, 154, 200, 229, 287, 314, and 345. Each identified cluster contains exactly 4 transactions initiated within a single 10-minute window, resulting in an aggregate exposure of ₹12,468 (averaging ₹1,558.50 per cluster attack). The uniform execution pattern—4 micro-transactions occurring in rapid succession within brief windows—is indicative of automated scripting, card testing, or credential stuffing attacks designed to exploit system response latencies before account locking triggers engage.

    Part C — Python payment reconciliation
        **Module Requirements & Scope**
        Part C requires developing an automated, reproducible Python reconciliation module (reconcile.py) to cross-evaluate internal transaction records (ledger.csv) against external settlement files from the payment gateway (gateway_export.csv).

        Functional Specifications:
            1. Implement a reusable function reconcile_payments(ledger_df, gateway_df) returning four DataFrames: 
                * Transactions missing in the gateway export 
                * Transactions missing in the ledger 
                * Amount mismatches  
                * Status mismatches  
            2. Perform missing record identification using Python. 
            3. Verify that empirical discrepancy counts match the noise injection rates (~5%, ~3%, ~2%, and ~2%) established during data synthesis. 

        **Detailed Architecture Logic (reconcile.py)**
            The architectural design of the payment reconciliation system is structured around a decoupled, two-stage data-processing pipeline that separates key existence checks from attribute comparisons to maximize efficiency. In the first stage, Python set operations calculate set differences between internal database keys and external gateway logs to rapidly flag un-captured transactions and untracked liabilities. Next, the pipeline performs an inner join via pd.merge() on the overlapping records, applying conditional masks to isolate numeric payment variances and status discrepancies into separate logging frameworks. By cleanly decoupling missing record detection from pairwise attribute matching, the system prevents redundant data scans, enforces third normal form relational compliance, and scales effectively across large transaction volumes.
        
        **Execution Results & Discrepancy Report**
            Executing reconcile_payments(ledger_df, gateway_df) against 547 ledger rows and 530 gateway export rows yields the following empirical counts:

            --- PAYMENT RECONCILIATION SUMMARY REPORT ---
            ================================================================
            Total Ledger Transactions: 547
            Total Gateway Export Transactions: 530

            1. Missing in Gateway Export     : 27 txns   | ₹26,673.00
            2. Missing in Ledger             : 10 txns   | ₹19,390.00
            3. Status Mismatches             : 9 txns    | ₹7,141.00
            4. Amount Mismatches (Abs Delta) : 16 txns   | ₹1,250.00
            ================================================================
            TOTAL MISMATCH TRANSACTIONS     : 62 txns
            TOTAL CAPITAL AT-RISK           : ₹54,454.00
            ================================================================

        Empirical vs. Injected Rate Alignment Table:
            | Category          | Observed Count | Observed Rate (%) | Injected Rate | Alignment |
            | Missing in Gateway|       27 txns  |             4.94% |           ~5% | Matched   |
            | Amount Mismatches |       16 txns  |             2.93% |           ~3% | Matched   |
            | Missing in Ledger |       10 txns  |             1.83% |           ~2% | Matched   |
            | Status Mismatches |        9 txns  |             1.65% |           ~2% | Matched   |
        
        ** Analysis**
            Category 1: Missing in Gateway Export (4.94% / 27 Transactions)
                A total of 27 transactions present in the internal ledger (totaling ₹26,673.00) are entirely absent from the gateway settlement file. This 4.94% rate directly matches the ~5% synthetic noise injection target. Operationally, this discrepancy indicates webhook delivery failures or dropped API acknowledgments where Paytm logged order creation locally, but the gateway failed to receive or process the underlying authorization. This represents direct liquidity risk, as uncollected customer funds remain un-settled by the gateway provider.

            Category 2: Missing in Ledger / Extra in Gateway (1.83% / 10 Transactions)
                There are 10 transactions recorded by the gateway (totaling ₹19,390.00) that possess no corresponding entry in Paytm’s internal database. Matching the ~2% target injection rate (1.83%), these unrecorded liabilities typically stem from high-concurrency database connection drops during order creation, or manual merchant refund overrides initiated directly on the gateway admin dashboard without triggering downstream database.

            Category 3: Amount Mismatches (2.93% / 16 Transactions)
                Evaluating common records surfaces 16 transactions exhibiting numerical variances between the ledger and gateway settlement figures (rate of 2.93%, aligning with the ~3% target). This structural consistency points to uncaptured fixed gateway processing fees, dynamic discount rate adjustments, or currency rounding differences, rather than arbitrary database corruption.

            Category 4: Status Mismatches (1.65% / 9 Transactions)
                Across overlapping records, 9 transactions exhibit status conflicts (1.65% observed rate vs. ~2% target). The predominant failure pattern involves transactions marked as captured in Paytm's internal ledger that are reported as failed within the gateway export log. This exposes the platform to immediate credit loss, as orders are fulfilled under the assumption of successful payment while the gateway processor has marked the transaction as rejected.


    Part D — Four-layer analytics dashboard (code-generated, not a live BI tool)
        **Module Requirements & Scope**
            Dashboard Format & Delivery:
                * Build a four-layer analytics dashboard using matplotlib. 
                * Deliver the dashboard as a set of saved chart images accompanied by written interpretations (no live Looker Studio or Power BI dependency is required). 
            
            Layer 1: Headline Layer (Scorecards)
                * Display 3–5 headline scorecard metrics as an image. 
                * Required metrics: Total GMV (in INR), Overall Success Rate, Reconciliation Match Rate, and Platform-Wide Chargeback Ratio. 
                * Strict Metric Definitions:
                    * match_rate = (count of transactions present in BOTH ledger.csv and gateway_export.csv with an identical amount_inr AND an identical status) / (total transaction count in ledger.csv). (Note: Amount mismatches, status mismatches, and rows missing in either file count as NOT matched for this headline scorecard metric). 
                    * chargeback_ratio (headline) = (count of transactions with status == "chargeback") / (count of all transactions), platform-wide, expressed as a percentage. 

            Layer 2: Trends Layer (Time Series)
                * Generate a time-series line/bar chart displaying Daily GMV alongside Daily Chargeback Count over the 30-day window. 
                
            Layer 3: Breakdown Layer (Bar Charts)
                * Generate bar charts illustrating GMV broken down by payment_method and GMV by merchant category (joined from merchants.csv). 
                
            Layer 4: Details Layer (Formatted Table Image)
                * Render a formatted table saved as an image (not a printed DataFrame output) displaying the Top 10 merchants by transaction count. 
                * Include conditional highlighting/flagging for any merchant whose per-merchant chargeback ratio exceeds 1.0%. 
                * Per-Merchant Metric Definition:
                    chargeback_ratio (per-merchant) = (count of that merchant's transactions with status == "chargeback") / (count of all of that merchant's transactions). 
            
            Written Interpretations:
                * Accompany each individual chart layer with a concise 2–4 sentence written analytical interpretation. 

        **Detailed Architecture (reconcile.py and layer1_headline_scorecards.png, layer2_trends_timeseries.png, layer3_breakdown_gmv.png, layer4_details_top10_merchants.png)**
            The architectural design of the four-layer analytics dashboard relies on an integrated, end-to-end data visualization pipeline built entirely using Python's pandas and matplotlib libraries to ensure complete platform independence from external business intelligence tools. Data ingestion initiates by joining relational transaction logs from ledger.csv with merchant data from merchants.csv and settlement records from gateway_export.csv to create a unified analytical dataset. The system then routes this dataset through four specialized rendering modules: Layer 1 aggregates high-level KPIs using custom-styled bounding shapes to display headline scorecards; Layer 2 constructs a dual-axis time-series chart mapping daily financial volume against chargeback frequency over time; Layer 3 builds side-by-side comparative bar charts detailing payment method adoption and merchant category breakdown; and Layer 4 leverages matplotlib.table to dynamically format a structured tabular overview of top merchants, applying conditional background fill logic to highlight entities exceeding acceptable chargeback thresholds. Each rendered figure is exported as a high-resolution image asset into a designated output directory, establishing an automated, code-generated reporting architecture.

        **Dashboard Visualizations & Written Interpretations**
            Layer 1: Headline Scorecards
                The platform processed an aggregate Gross Merchandise Value (GMV) of ₹382,603 across 547 ledger transactions, maintaining an overall transaction processing success rate of 85.6%. Strict reconciliation against payment gateway settlement logs reveals a Reconciled Match Rate of 90.5% (495 fully matching transactions), indicating that 9.5% of internal records suffer from status, amount, or existence discrepancies. Meanwhile, the platform-wide chargeback ratio stands at 5.12% (28 total dispute cases), significantly exceeding standard card network thresholds (1.0%) and signaling high operational exposure to fraudulent activity.

            Layer 2: Trends Layer (Time Series)
                Over the 30-day monitoring window, daily GMV exhibits dynamic cyclicality with revenue spikes exceeding ₹28,000 on high-volume processing days. Chargeback occurrences show periodic clustering, spiking to an all-time high of 4 disputes in a single day towards the final week of the month. Notably, chargeback spikes frequently lag major GMV surges by 2 to 4 days, reflecting the operational latency of cardholder dispute filings following velocity testing or fraud events.

            Layer 3: Breakdown Layer (Bar Charts)
                Payment method volume is dominated by UPI, which accounts for ₹172.3k (45.0%) of total GMV, followed by Card payments at ₹102.4k and Wallets at ₹71.3k. Across merchant verticals, ecommerce (₹79.9k), travel (₹75.2k), and grocery (₹71.9k) represent the primary revenue drivers, collectively contributing over 59% of processed GMV. Lower-volume categories such as bill_payment (₹26.3k) and recharge (₹15.1k) exhibit lower ticket sizes, reflecting routine micro-transaction behavior.

            Layer 4: Details Layer (Merchant Risk Performance Table)
                Examining the top 10 merchants by transaction volume reveals significant risk concentration, with 7 out of 10 merchants flagged as HIGH RISK for exceeding the 1.0% per-merchant chargeback ratio threshold. Merchant_027 and Merchant_029 exhibit severe risk profiles with chargeback ratios of 18.75% and 15.79% respectively, driven by 3 disputes each. Conversely, high-volume entities like Merchant_016 (20 txns, ₹11,130 GMV) and Merchant_009 (18 txns, ₹4,982 GMV) maintained a 0.00% dispute rate, proving that elevated volume does not inherently correlate with fraud exposure.        

**===================================================================================================**
                Part 2 — Credit Risk & Lending ML (/credit_risk_lending_ml)
**===================================================================================================**
    
    Part A — EDA and preprocessing
        **Module Requirements & Scope**
            Part A of the Machine Learning pipeline requires preparing the raw credit applicant dataset (credit_applicants.csv) for predictive modeling through exploratory data analysis, feature engineering, statistical imputation, categorical encoding, and feature scaling:
                * Load credit_applicants.csv and compute the exact platform-wide default rate and the missing percentage for credit_bureau_score. 
                * Create a binary indicator flag (is_thin_file) set to 1 when credit_bureau_score is missing and 0 otherwise, preserving new-to-credit applicants without dropping records or leaking fitted statistics prior to splitting. 
                * Perform a 75/25 stratified train/test split (random_state=42) based on the target variable (default). 
                * Compute the median credit_bureau_score strictly from the non-missing observations within the training split, and apply this exact training-derived median value to fill missing values across both training and testing splits. 
                * State and execute an appropriate encoding strategy (One-Hot or label encoding) for the nominal feature employment_type. 
                * Standardize all numerical features using StandardScaler, fitting the scaler parameters strictly on the training split.  

        **Detailed Architecture Logic (credit_risk_lending_ml_generate_data.py)**
            The data preprocessing pipeline isolates training data statistics from the test evaluation set to prevent data leakage and maintain model governance. In the first stage, raw credit records are ingested into pandas, and an is_thin_file binary flag is created to identify missing data before splitting. Next, the pipeline performs a stratified train/test split based on the default target label to keep class proportions consistent across both sets. Finally, key transformation metrics—such as median bureau scores, standard scaling parameters (means and standard deviations), and one-hot encoded categorical columns—are calculated using only the training split before being applied across both datasets to generate the final numerical feature matrices.
        
        **Output and Analysis**
            Exploratory Metrics & Preprocessing Execution Report

            --- EDA AND PREPROCESSING SUMMARY REPORT ---
                ================================================================
                Total Applicant Records            : 400
                Exact Measured Default Rate        : 20.25% (81 / 400 applicants)
                Missing Credit Bureau Score Rate   : 20.00% (80 / 400 applicants)
                ================================================================
                Train/Test Split Ratio             : 75 / 25 (Stratified, random_state=42)
                Training Set Size                  : 300 applicants (61 defaults | 20.33%)
                Testing Set Size                   : 100 applicants (20 defaults | 20.00%)
                ================================================================
                Training-Derived Bureau Median     : 612.00
                Categorical Encoding Strategy      : One-Hot Encoding (drop_first=True)
                Numerical Scaling Strategy         : StandardScaler (fit on train only)
                ================================================================

        Operational Analysis Narrative:
            Measured Metrics & Initial Exploration: 
                * Exact Measured Default Rate: 20.25% (81 defaults out of 400 total applicants). 
                * Missing Credit Bureau Score Rate: 20.00% (80 missing values out of 400 total applicants).

            Feature Engineering: Thin-File Flag
                To accommodate new-to-credit applicants without dropping them from the pipeline, a binary indicator is_thin_file was created: is_thin_file =  1 if credit_bureau_score is NaN otherwise 0

                Justification: Exploration of credit_applicants.csv reveals a platform default rate of 20.25% and a missing credit_bureau_score rate of 20.00% (80 out of 400 applicants). Rather than discarding these missing rows—which would completely exclude new-to-credit applicants and undermine alternative-data credit scoring objectives—the system engineers the binary flag is_thin_file directly from raw availability. Because is_thin_file relies solely on element-wise checks (isna()), it is computed prior to splitting without risking statistical data leakage.

            Train/Test Split & Stratification Choice:
                The dataset was partitioned into a 75% Training Split (300 applicants) and a 25% Test Split (100 applicants) using random_state=42 and stratify=y. Stratifying on default enforces class distribution parity across both splits. 
                    * Training Split (300 rows): 61 defaults (20.33% default rate). 
                    * Test Split (100 rows): 20 defaults (20.00% default rate). 
                Without stratification, a random split on a small sample ($N=400$) could introduce sampling bias, resulting in severe class imbalance drift between model training and evaluation.
            
            Alternate-Data-Driven Imputation Choice
                Missing values in credit_bureau_score were filled using the training-derived median score of 612.00, computed exclusively from non-missing training observations.
                
                Justification:
                    Missing credit_bureau_score values are filled using the training-derived median of 612.00. This median is calculated exclusively from non-missing training observations and mapped to both training and testing splits. Imputing with the training median provides a neutral baseline for thin-file applicants while preventing test-set information from leaking into model fitting. When combined with the is_thin_file indicator, downstream classifiers can distinguish traditional applicants from thin-file applicants, leveraging alternative features (such as upi_monthly_inflow_inr) to assess credit risk.
            
            Categorical Encoding Choice
                One-Hot Encoding was applied to employment_type using pd.get_dummies(..., drop_first=True).
                    * Base Reference Category: gig
                    * Encoded Feature Columns: employment_type_salaried, employment_type_self_employed
                
                Justification: employment_type is nominal (salaried, self_employed, gig) with no inherent order. One-Hot Encoding with drop_first=True transforms this variable into two binary features (employment_type_salaried and employment_type_self_employed), using gig as the baseline to avoid multicollinearity.

            Feature Scaling
                Numerical features (age, monthly_income_inr, existing_loans_count, credit_utilization_ratio, upi_monthly_inflow_inr, bounced_payments_count, credit_bureau_score) were standardized using StandardScaler.
                    * scaler.fit_transform(X_train[numeric_cols]) was executed solely on the training split. 
                    * scaler.transform(X_test[numeric_cols]) was executed on the test split using training-derived parameters. 

    Part B — Classification models
        **Module Requirements & Scope**
            Part B of the Machine Learning pipeline focuses on training, evaluating, and applying supervised credit default classification models:
                * Train a baseline LogisticRegression(random_state=42) model and a DecisionTreeClassifier(random_state=42) model using the identical stratified 75/25 preprocessed train/test datasets. 
                * Evaluate both models across standard metrics—Confusion Matrix (TN, FP, FN, TP), Accuracy, Precision, Recall, F1-Score, and ROC-AUC—presenting the metrics side-by-side in a comparative evaluation table. 
                * Extract predicted default probabilities from Logistic Regression, bucket applicants into 4 distinct risk tiers (quartiles), map illustrative interest rate ranges to each tier (lower risk --> lower rate), and evaluate actual observed default rates across tiers to check monotonicity. 

        **Detailed Architecture Logic (credit_risk_lending_ml_generate_data.py)**
            The classification and risk-based pricing architecture ingests the standard feature matrices developed in Part A. Model fitting is executed concurrently on the preprocessed training set. Post-training, both models evaluate the held-out test split to generate hard class predictions (y) and continuous default probability scores (p). Performance metrics are computed via a standardized metric that compiles confusion matrix components alongside threshold-dependent and threshold-independent performance statistics into a structured Markdown comparison table. For credit risk stratification, the continuous probabilities generated by Logistic Regression are routed to a quartile binning engine (pd.qcut), partitioning the test set into four equal-sized risk buckets. The system computes actual empirical default rates per bucket to verify monotonic risk ordering before mapping corresponding interest rate ranges, establishing an operational risk-based pricing framework.

        **Model Output and Evaluation**
            Output:
            =========================================================================
             MODEL EVALUATION COMPARISON TABLE (TEST SET)                
            =========================================================================
                                    Confusion Matrix (TN, FP, FN, TP)  Accuracy  Precision  Recall  F1-Score  ROC-AUC
            Logistic Regression                    [[69, 11], [13, 7]]      0.76   0.388889    0.35  0.368421  0.71875
            Decision Tree Classifier               [[59, 21], [14, 6]]      0.65   0.222222    0.30  0.255319  0.51875
            =========================================================================
            ==========================================================================
                            RISK-BASED PRICING TABLE (TEST SET)                    
            ==========================================================================
                risk_tier     total_applicants pred_prob_range  observed_defaults  observed_default_rate   interest_rate_range
            Tier 1 (Low Risk)            25    0.5% - 3.6%              2                    8.0               9.5% - 12.0%
            Tier 2 (Medium-Low Risk)     25    3.6% - 14.6%             3                   12.0              13.0% - 16.5%
            Tier 3 (Medium-High Risk)    25   15.2% - 34.3%             5                   20.0              17.0% - 21.0%
            Tier 4 (High Risk)           25   35.9% - 94.6%             10                  40.0      22.0% - 28.0% (or Decline)
            ==========================================================================

        **Analysis**
            Comparative Model Performance Analysis
                Logistic Regression significantly outperforms the unconstrained Decision Tree Classifier across all key evaluation metrics:
                    * Predictive Discrimination (ROC-AUC): Logistic Regression achieves an AUC of 0.7188, indicating strong signal separation between defaulting and non-defaulting applicants. In contrast, the unconstrained Decision Tree achieves an AUC of 0.5188 (near random chance), suffering from severe overfitting on the small training set. 
                    * Classification Accuracy & F1-Score: Logistic Regression yields 76.00% Accuracy and an F1-Score of 0.3684, compared to the Decision Tree's 65.00% Accuracy and 0.2553 F1-Score. 
                    * False Positive Control: Decision Tree generates 21 False Positives (crediting high-risk applicants who default), resulting in a low Precision of 22.22%, whereas Logistic Regression limits False Positives to 11 (38.89% Precision). 
            
            Monotonicity Verification of Risk-Based Pricing
            The risk-based pricing engine demonstrates strict monotonicity across all four tiers:
                * Tier 1 (Lowest Risk): Applicants with predicted probabilities between 0.5% and 3.6% exhibit an actual default rate of 8.00% and receive prime lending rates (9.5% – 12.0%). 
                * Tier 2 (Medium-Low Risk): Applicants with probabilities between 3.6% and 14.6% show a 12.00% default rate. 
                * Tier 3 (Medium-High Risk): Applicants with probabilities between 15.2% and 34.3% show a 20.00% default rate. 
                * Tier 4 (Highest Risk): Applicants with probabilities between 35.9% and 94.6% suffer an observed default rate of 40.00%, justifying higher risk premiums (22.0% – 28.0%) or outright loan rejections. 
            This monotonic progression confirms that Logistic Regression probabilities reliably rank-order credit risk, enabling risk-adjusted pricing that protects credit margins while maintaining competitive pricing for low-risk applicants.

    Part C — Anomaly detection and optional segmentation
        **Module Requirements & Scope**
        Part C of the Machine Learning pipeline focuses on unsupervised learning techniques applied to behavioral transactions and credit applicant profiling:
            * Load txn_behaviour.csv and isolate the three numeric behavioral features: txn_hour, is_new_device, and txn_amount_inr. 
            * Standardize the numerical features using StandardScaler. 
            * Train scikit-learn's IsolationForest(random_state=42, contamination=...) with a contamination rate matching the exact injected anomaly proportion (15 / 265 ≈ 5.7%). 
            * Evaluate prediction flags (anomaly_pred = -1) against injected ground truth anomalies (transaction IDs starting with BTXNA) and report the exact recall percentage. 
            * Perform K-Means clustering on the standardized credit applicant dataset (credit_applicants.csv), using the Calinski-Harabasz index or Elbow method to determine optimal k, and evaluate whether any cluster over-indexes on the default label.

        **Detailed Architecture Logic (credit_risk_lending_ml_generate_data.py)**
        The unsupervised learning pipeline consists of two modular workflows:
            1. Transaction Anomaly Detection Subsystem: Reads transaction records from txn_behaviour.csv, standardizes numerical transaction attributes, and feeds the feature matrix into an IsolationForest estimator. The model isolates anomalous records by randomly splitting features; shorter path lengths in isolation trees correspond to high-risk anomalous behavior. Predictions are mapped back to ground truth flags (BTXNA prefix) to generate a recall metric. 
            2. Applicant Segmentation Subsystem: Ingests preprocessed credit applicant features, standardizes numerical columns, and evaluates cluster separation across k using the Calinski-Harabasz Index and Inertia Elbow methods. Upon selecting optimal k value, a final KMeans model groups applicants into behavioral segments. The pipeline aggregates group statistics to identify high-risk segments over-indexing on credit defaults. 
        
        **Output and Analysis**
            Anomaly Detection Performance & Recall Report
                The Isolation Forest model evaluated 265 transaction records with a contamination parameter set to (15 / 265 ≈ 5.7%):
                   ==========================================================================
                   ISOLATION FOREST ANOMALY DETECTION REPORT                    
                   ==========================================================================
                   Total Behavioral Transactions Evaluated : 265
                   Total Injected Seeded Anomalies (BTXNA) : 15
                   Contamination Parameter                 : 0.056604 (15 / 265)
                   ================================================================
                   Flagged Seeded Anomalies                : 11
                   Missed Seeded Anomalies                 : 4
                   Isolation Forest Anomaly Recall         : 73.33%
            
            K-Means:
                ==========================================================================================
                    K-MEANS CLUSTERING REPORT                                              
                ==========================================================================================
                Calinski-Harabasz Scores for k=[2, 3, 4, 5, 6]: ['49.07', '43.30', '40.74', '39.28', '37.32']
                Optimal K selected via Calinski-Harabasz Index: 2
                ==========================================================================================
                cluster  total_applicants  observed_defaults  default_rate_pct   avg_income  avg_utilization  avg_bounced_payments  thin_file_pct
                    0         195         55       28.205128      74525.010256      0.536667       1.389744        23.589744
                    1         205         26       12.682927      85991.004878      0.477415       1.053659        16.585366
                ==========================================================================================
        
            Cluster quality was evaluated across k=2 to k=6 using the Calinski-Harabasz Index and the Elbow Method (Inertia). The Calinski-Harabasz score peaked at k=2 (score = 49.07), indicating maximum cluster separation.

        Analysis of Isolation Forest Anomaly Detection
            Setting the contamination parameter to 5.66% allowed the IsolationForest to isolate 11 out of the 15 injected anomalies, achieving a 73.33% Recall rate against ground truth.
                * The flagged transactions consistently represented high-risk patterns: off-hour execution (1:00 AM – 4:00 AM), new device access (is_new_device = 1), and high transaction amounts (₹14,999 – ₹24,999). 
                * The 4 unflagged anomalies fell closer to standard variance boundaries in amount or hour, highlighting the need to combine isolation trees with rule-based thresholding in live fraud monitoring pipelines. 

        Analysis of Cluster Over-Indexing on Credit Default
            The K-Means clustering (k=2) revealed distinct behavioral segments with noticeable risk divergence:
                * Cluster 0 exhibits an observed default rate of 28.21% (55 defaults out of 195 applicants), which is 2.22 times higher than Cluster 1 (12.68%) and significantly above the platform baseline default rate of 20.25%. 
                Profile: Applicants in Cluster 0 display lower average monthly income (₹74,525), higher credit utilization (53.67%), more bounced payments (1.39 average), and a higher proportion of thin-file applicants (23.59%). 
                
                * Cluster 1 records an observed default rate of 12.68% (26 defaults out of 205 applicants). 
                Profile: Applicants feature higher monthly income (₹85,991), lower credit utilization (47.74%), fewer bounced payments (1.05 average), and fewer thin-file profiles (16.59%). 
            This segmentation confirms that combining alternative transaction signals (bounced payments, income, utilization) effectively isolates high-risk credit pools even prior to supervised model scoring.   

    Part D — Bias-awareness note and final recommendation
        **Module Requirements & Scope**
        Part D requires synthesizing ethical governance and model selection for deployment readiness:
            * Write a concise (200–400 word) governance note examining how non-explicit features (employment_type, monthly_income_inr, credit_bureau_score) can act as proxies for protected attributes (e.g., gender, age, socioeconomic background) in production credit scoring models, and recommend operational safeguards such as human-in-the-loop review. 
            * Consolidate evaluation metrics across all machine learning tasks, bringing together both classifiers' performance metrics (Task 4/Part B) and the Isolation Forest anomaly detection recall (Task 6/Part C). 
            * Provide a clear 3–5 sentence deployment recommendation for Paytm Postpaid, referencing exact metric values to justify model choice.

        **Detailed Architecture **
            The governance framework links algorithmic risk scoring to human oversight mechanisms. Although explicit protected attributes are excluded from feature engineering, input variables pass through an audit layer that monitors demographic parity across employment categories and income tiers. In production, model output route through a threshold gate: high-confidence approvals or declines are processed automatically, while border or thin-file declines trigger a human-in-the-loop maker-checker review workflow. The final selection framework prioritizes metric stability, non-linear calibration, and operational risk mitigation to choose the optimal production pipeline.

        **Interpretation and Analysis**
            Proxy Discrimination & Disparity Risks:
            Even in the absence of explicit demographic variables like gender, ethnicity, or region, ML credit scoring models remain susceptible to proxy discrimination. In real-world deployments, non-explicit features frequently encode systemic societal biases:
                * employment_type: Informal or gig economy workers are disproportionately female or from historically underrepresented socioeconomic groups who lack access to formal salaried positions. Penalizing gig or self-employed categories directly reduces credit access for these demographics. 
                * monthly_income_inr: Gender pay gaps and regional income disparities mean raw income thresholds can disproportionately lower credit scores for female applicants and rural populations, regardless of individual debt repayment reliability. 
                * credit_bureau_score: Traditional bureau coverage inherently favors older, urban, higher-income demographics. Younger applicants and first-time borrowers face a systematic thin-file penalty when bureau scores are missing or sparse. 
            
            Recommended Governance Framework: Maker-Checker Human-in-the-Loop
            To mitigate algorithmic bias and prevent unfair credit exclusion, the deployment architecture must establish a Maker-Checker Human-in-the-Loop (HITL) review policy:
                1. Automated rejections for thin-file applicants (is_thin_file == 1) or applicants flagged in Tier 4 high-risk buckets must not be final. Instead, they are routed to credit underwriters for secondary evaluation. 
                2. Underwriters examine alternative cash-flow signals (e.g., upi_monthly_inflow_inr, utility payment consistency) to evaluate creditworthiness independently of traditional credit bureau scores. 
                3. Risk teams must track Demographic Parity Ratio (DPR) and Equal Opportunity Difference (EOD) across employment types and gender/location proxies, re-calibrating score thresholds whenever approval rate disparities exceed the 80% rule. 

        Final Unified Model Performance Table
        The summary table below aggregates evaluation metrics across all machine learning tasks:
        | Model  Component    | Target Task       | Primary Evaluated Metric(s)   | Observed Metric Value    | Performance Assessment |
        | Logistic Regression | Default Prediction | ROC-AUC / Accuracy / F1      | 0.7188 / 76.00% / 0.3684 | Selected (Strong discrimination & linear stability) |
        | Decision Tree Classifier | Default Prediction | ROC-AUC / Accuracy / F1 | 0.5188 / 65.00% / 0.2553 | Rejected (Severe overfitting & low generalization)  |
        | Isolation Forest | Anomaly Detection | Seeded Anomaly Recall | 73.33% (11/15 flagged) | Approved for Fraud Screening (Effective detection of high-risk transactions) |

        Final Model Deployment Recommendation
            We recommend deploying the Logistic Regression model for Paytm Postpaid credit scoring, alongside the Isolation Forest pipeline for real-time transaction anomaly detection. Logistic Regression achieves a strong ROC-AUC of 0.7188 and 76.00% Accuracy, outperforming the unconstrained Decision Tree (0.5188 ROC-AUC and 65.00% Accuracy), which degrades to near-random guessing. Furthermore, Logistic Regression generates well-calibrated, monotonic risk tier probabilities—scaling observed default rates strictly from 8.00% in Tier 1 to 40.00% in Tier 4—enabling precise risk-based pricing. Operating in tandem, the Isolation Forest model secures transaction monitoring by capturing 73.33% of seeded fraud anomalies.
            
**===================================================================================================**
      Part 3 — AI-Augmented FinTech Advisory & Blockchain Risk (/ai_advisory_blockchain)
**===================================================================================================**
    
    Part A — Portfolio advisory agent (agentic think-act-observe pattern)
        **Module Requirements & Scope**
        1. The portfolio advisory agent follows the explicit Think-Act-Observe pattern to deliver risk-aware portfolio recommendations:
            1. Think Stage: The agent reads the investor profile (investor_id, risk_tolerance) and maps it to the prescribed equal-weight (1/3) allocation matrix: 
                * Conservative: {"PAYBOND", "PAYGOLD", "PAYRETAIL"}
                * Moderate: {"PAYRETAIL", "PAYINFRA", "PAYGOLD"}
                * Aggressive: {"PAYTECH", "PAYFIN", "PAYINFRA"}
            2. Act (tool call): call a get_stock_data(ticker) "tool" function that looks up beta/analyst_expected_return/std_dev from STOCK_UNIVERSE for each ticker in the prescribed allocation (this simulates an external-API tool call; no real API is needed since the data is local). 
            3. Observe → decide: using the prescribed 1/3-each allocation for the investor's risk_tolerance tier, compute the portfolio's CAPM-expected return (per stock, E(R) = R_f + β(E(R_m) − R_f), using ONLY beta — never analyst_expected_return — then weight-averaged across the 3 tickers) and portfolio variance using: Var(R_p) = Σᵢ wᵢ²σᵢ² + 2·Σ_{i < j} wᵢwⱼ·Cov(Rᵢ,Rⱼ), with Cov(Rᵢ,Rⱼ) = ρ·σᵢ·σⱼ and a stated pairwise correlation ρ = 0.3 for every pair of the three tickers in the prescribed allocation. Convert variance to portfolio standard deviation.
            4. Human-in-the-loop escalation: if the computed portfolio standard deviation exceeds 20%, do not auto-finalize the recommendation — instead print/return an "ESCALATED_TO_HUMAN_ADVISOR" flag with the computed numbers attached. Otherwise, finalize the recommendation. With the prescribed allocation table and ρ = 0.3, the expected pattern is deterministic: Conservative (INV01, ~8.44% std dev) and Moderate (INV02, INV04, ~12.57% std dev) must NOT escalate; Aggressive (INV03, INV05, ~20.58% std dev) must escalate. 
        2. The final narrative sentence describing the recommendation is the only part gated by MOCK_LLM. Mock mode (graded baseline): build the sentence from an f-string template inserting the computed numbers (e.g., f"For {risk_tolerance} investor {investor_id}, we recommend an allocation across {tickers} with an expected portfolio return of {return:.1%} and volatility of {vol:.1%}."). Optional MOCK_LLM=0 extension: prompt the LLM to phrase the same numbers more naturally. Run all 5 investor profiles and record each result.

        **Detailed Architecture Logic (advisory_agent.py) **
            The architectural design of the Agentic Portfolio Advisory Engine relies on a deterministic **Think-Act-Observe** control loop that integrates quantitative portfolio theory with automated safety governance. The workflow begins with the **Think** stage, where the agent ingests an investor profile and evaluates its risk tolerance against a strict, rule-based lookup table. Instead of allowing an unconstrained language model to guess allocations, the system deterministically maps Conservative profiles to equal parts `PAYBOND`, `PAYGOLD`, and `PAYRETAIL`; Moderate profiles to `PAYRETAIL`, `PAYINFRA`, and `PAYGOLD`; and Aggressive profiles to `PAYTECH`, `PAYFIN`, and `PAYINFRA`. This design ensures regulatory compliance, eliminates model hallucination at the allocation stage, and establishes a predictable foundation for downstream financial modeling.

            Transitioning to the **Act** stage, the agent invokes an internal tool function, `get_stock_data(ticker)`, to fetch systematic risk (beta) and asset volatility (sigma) metrics for each ticker in the chosen allocation. In the subsequent **Observe** stage, the agent executes double-precision mathematical analytics: it calculates individual expected returns purely via the Capital Asset Pricing Model (E(R_i) = R_f + \beta_i(E(R_m) - R_f))—intentionally ignoring subjective analyst estimates—and combines them into a weighted portfolio return. It then computes total portfolio variance by combining individual weighted asset variances with cross-asset covariances using a fixed pairwise correlation (ρ = 0.3), ultimately deriving the annualized portfolio standard deviation (σp).

            To safeguard capital, the architecture inserts an automated **Human-in-the-Loop (HITL)** risk gate immediately following the Observe calculations. The engine evaluates portfolio volatility against a hard 20.0% threshold: portfolios exceeding this limit (such as the high-beta Aggressive profiles `INV03` and `INV05` at approximately 20.58% volatility) are flagged as `ESCALATED_TO_HUMAN_ADVISOR`, halting automated execution for manual review. Portfolios at or below the threshold (such as Conservative `INV01` at approx 8.44% and Moderate `INV02`/`INV04` at approx 12.57%) are marked `AUTO_APPROVED`. Finally, narrative generation is isolated behind a `MOCK_LLM` environment gate. Under the default mock setting, a rule-based template inserts the exact calculated return and volatility metrics into a deterministic string, while an optional `MOCK_LLM=0` setting forwards the pre-computed metrics to an external API purely for natural language formatting, maintaining complete separation between quantitative calculations and narrative generation.

        **Execution and Interpretation**
        Execution Results Summary Across 5 Investor Profiles
        | Investor ID | Risk        | Prescribed Allocation        | CAPM Return | Portfolio Volatility | Status    | Escalted? |
        | INV01       | Conservative| PAYBOND, PAYGOLD, PAYRETAIL  |       9.20% |                8.44% | FINALIZED | No        |
        | INV02       | Moderate    | PAYRETAIL, PAYINFRA, PAYGOLD |      11.30% |               12.57% | FINALIZED | No        |
        | INV03       | Aggressive  | PAYTECH, PAYFIN, PAYINFRA    |      15.00% |               20.58% | ESCALATED | Yes       |
        | INV04       | Moderate    | PAYRETAIL, PAYINFRA, PAYGOLD |      11.30% |               12.57% | FINALIZED | No        |
        | INV05       | Aggressive  | PAYTECH, PAYFIN, PAYINFRA    |      15.00% |               20.58% | ESCALATED | Yes       |

        Detailed Output Logs & Generated Narratives
            [INV01] Risk Tier: Conservative
            Allocated Tickers: ['PAYBOND', 'PAYGOLD', 'PAYRETAIL']
            CAPM Return E(R) : 9.20%
            Portfolio Risk σ : 8.44%
            Execution Status : FINALIZED
            Narrative Output : For Conservative investor INV01, we recommend an allocation across ['PAYBOND', 'PAYGOLD', 'PAYRETAIL'] with an expected portfolio return of 9.2% and volatility of 8.4%.

            [INV02] Risk Tier: Moderate
            Allocated Tickers: ['PAYRETAIL', 'PAYINFRA', 'PAYGOLD']
            CAPM Return E(R) : 11.30%
            Portfolio Risk σ : 12.57%
            Execution Status : FINALIZED
            Narrative Output : For Moderate investor INV02, we recommend an allocation across ['PAYRETAIL', 'PAYINFRA', 'PAYGOLD'] with an expected portfolio return of 11.3% and volatility of 12.6%.

            [INV03] Risk Tier: Aggressive
            Allocated Tickers: ['PAYTECH', 'PAYFIN', 'PAYINFRA']
            CAPM Return E(R) : 15.00%
            Portfolio Risk σ : 20.58%
            Execution Status : ESCALATED_TO_HUMAN_ADVISOR
            Narrative Output : For Aggressive investor INV03, we recommend an allocation across ['PAYTECH', 'PAYFIN', 'PAYINFRA'] with an expected portfolio return of 15.0% and volatility of 20.6%.

            [INV04] Risk Tier: Moderate
            Allocated Tickers: ['PAYRETAIL', 'PAYINFRA', 'PAYGOLD']
            CAPM Return E(R) : 11.30%
            Portfolio Risk σ : 12.57%
            Execution Status : FINALIZED
            Narrative Output : For Moderate investor INV04, we recommend an allocation across ['PAYRETAIL', 'PAYINFRA', 'PAYGOLD'] with an expected portfolio return of 11.3% and volatility of 12.6%.

            [INV05] Risk Tier: Aggressive
            Allocated Tickers: ['PAYTECH', 'PAYFIN', 'PAYINFRA']
            CAPM Return E(R) : 15.00%
            Portfolio Risk σ : 20.58%
            Execution Status : ESCALATED_TO_HUMAN_ADVISOR
            Narrative Output : For Aggressive investor INV05, we recommend an allocation across ['PAYTECH', 'PAYFIN', 'PAYINFRA'] with an expected portfolio return of 15.0% and volatility of 20.6%.

        **Analysis**
        The execution results across the five investor profiles demonstrate how the advisory engine successfully balances portfolio risk-return optimization with deterministic automated risk governance. Expected returns scale monotonically alongside risk tolerance: the Conservative allocation (`INV01`) yields an expected CAPM return of 9.20% at an annualized portfolio volatility of 8.44%, the Moderate allocation (`INV02` and `INV04`) achieves an 11.30% return at 12.57% volatility, and the Aggressive allocation (`INV03` and `INV05`) delivers a 15.00% return at 20.58% volatility. The engine's automated Human-in-the-Loop (HITL) safety gate functions precisely at the designated sigma_p > 20.00% threshold—allowing lower-risk Conservative and Moderate profiles to pass through as `FINALIZED` (Auto-Approved) while flagging high-beta Aggressive profiles as `ESCALATED`, successfully halting automated trade execution over the risk boundary to mandate human advisor intervention.    
        
    Part B — Structured disclosure extraction

        **Module Requirements & Scope**
        Here is the simplified bullet-point breakdown of the task requirements:

            * Create `extract_disclosure.py` with a function named `extract_signals(snippet: str)`.
            * The function must return a dictionary containing three specific keys:
                - `risk_flags`: A list of identified risk categories (e.g., `["litigation exposure"]`).
                - `hedging_detected`: `True` if conditional or uncertain language is present, otherwise `False`.
                - `sentiment`: Must be strictly one of `"confident"`, `"cautious"`, or `"neutral"`.
            * Rule-Based Mock Logic :
                - Risk Flags: Flag `"litigation"` as litigation risk, `"regulatory"` as regulatory risk, and `"customer concentration risk"` .
                - Hedging: Set `hedging_detected = True` if the text contains `"assuming"`, `"cautiously"`, or `"visibility"`.
                - Sentiment: Classify as `"confident"` if the text contains `"confident"` or `"approved"`, as `"cautious"` if any hedging phrase is present, or default to `"neutral"`.
            * Optional LLM Extension (`MOCK_LLM=0`): If using an external LLM, parse the output into the required JSON structure.
            * Testing Scope: Execute the function across all **6 disclosure snippets** (`doc_01` to `doc_06`) and display/record the results.

        **Detailed Architecture design (extract_disclosure.py)**
        The disclosure extraction module (extract_disclosure.py) parses corporate disclosure snippets to output structured JSON risk and sentiment metadata:
            1. Rule-Based Mock Extractor (extract_signals_mock):
                * Risk Flags: Uses regex/keyword matching to identify litigation ("litigation"), regulatory exposure ("regulatory", "compliance"), and customer concentration ("customer", "revenue", or percentage phrases like "42 percent"). 
                * Hedging Detection: Scans for operational hedging keywords ("assuming", "cautiously", or "visibility"). 
                * Sentiment Classification:
                    - Sets sentiment = "confident" if the text contains "confident" or "approved". 
                    - Sets sentiment = "cautious" if hedging is detected. 
                    - Defaults to sentiment = "neutral" otherwise. 
            2. Schema Validation & Fallback Mechanism:
                * Validates target structure: risk_flags (list), hedging_detected (bool), and sentiment {"confident"}{"cautious"}, {"neutral"}. 
                * On MOCK_LLM=0, the function sends structured prompts to Groq API with json_object enforcement, attempting up to 2 LLM completion calls before gracefully falling back to the deterministic mock logic if schema validation fails. 
        
        **Interpretation and Analysis**
            Summary Table across All 6 Disclosure Documents:
            | Document ID | Extracted Risk Flags            | Hedging Detected | Sentiment | Primary Keyword Triggers    |
            | doc_01      | []                              | True             | cautious  | "assuming"                  |
            | doc_02      | ['litigation exposure']         | False            | neutral   | "litigation"                |
            | doc_03      | ['customer concentration risk'] | False            | neutral   | "42 percent", "account for" |
            | doc_04      | []                              | True             | cautious  | "cautiously", "visibility"  |
            | doc_05      | []                              | False            | confident | "confident", "approved"     |
            | doc_06      | ['regulatory exposure']         | False            | neutral   | "regulatory", "compliance"  |

        Formatted Execution Log
            =========================================================================================
                            DISCLOSURE STRUCTURED SIGNAL EXTRACTION RESULTS                         
            =========================================================================================
            [doc_01]
            Snippet          : doc_01: Assuming input costs remain stable through the next two quarters, we expect margins to hold at current levels.
            Risk Flags       : []
            Hedging Detected : True
            Sentiment        : 'cautious'
            -----------------------------------------------------------------------------------------
            [doc_02]
            Snippet          : doc_02: The company faces an ongoing litigation matter related to a former vendor contract; management believes the exposure is not material.
            Risk Flags       : ['litigation exposure']
            Hedging Detected : False
            Sentiment        : 'neutral'
            -----------------------------------------------------------------------------------------
            [doc_03]
            Snippet          : doc_03: Our top three customers together account for approximately 42 percent of total revenue this year.
            Risk Flags       : ['customer concentration risk']
            Hedging Detected : False
            Sentiment        : 'neutral'
            -----------------------------------------------------------------------------------------
            [doc_04]
            Snippet          : doc_04: We remain cautiously optimistic about demand recovery, though visibility beyond the next quarter is limited given macro uncertainty.
            Risk Flags       : []
            Hedging Detected : True
            Sentiment        : 'cautious'
            -----------------------------------------------------------------------------------------
            [doc_05]
            Snippet          : doc_05: The board is confident in the long-term strategy and has approved an expanded capital expenditure plan for the coming year.
            Risk Flags       : []
            Hedging Detected : False
            Sentiment        : 'confident'
            -----------------------------------------------------------------------------------------
            [doc_06]
            Snippet          : doc_06: A recent regulatory notice has been received regarding data-localization compliance; the company is in active dialogue with the regulator.
            Risk Flags       : ['regulatory exposure']
            Hedging Detected : False
            Sentiment        : 'neutral'
            -----------------------------------------------------------------------------------------

        Analysis and Interpretation
            1. Risk Identification Accuracy:
                * doc_02 (Litigation Exposure): The model accurately identifies active legal proceedings related to vendor contracts. Although management deems the financial exposure non-material, flagging litigation risk ensures transparency for compliance oversight.
                * doc_03 (Customer Concentration Risk): The rule-based engine successfully captures revenue dependence where three top customers generate 42% of total revenue, signaling high client-concentration risk.
                * doc_06 (Regulatory Exposure): Successfully flags compliance mandates and regulatory interactions triggered by data-localization notices.
            
            2. Hedging & Sentiment Dynamics:
                * Hedging Triggers (doc_01 & doc_04): Disclosures containing conditional management assertions (e.g., "Assuming input costs remain stable", "cautiously optimistic", "limited visibility") are correctly flagged with hedging_detected = True and assigned a cautious sentiment stance.
                * Positive Outlook (doc_05): Expresses explicit board confidence and approved capital expenditure plans, mapping directly to a confident sentiment rating with zero operational risk flags.
                * Neutral Stance (doc_02, doc_03, doc_06): Standard operational and factual disclosures lacking explicit growth optimism or macro hedging language default appropriately to neutral sentiment.
            
            3. Production Implementation Strengths:
                * Schema Strictness: Enforces explicit type checking across key parameters (list, bool, and categorical sentiment strings), guaranteeing consistent JSON payloads.
                * LLM Resilience: On MOCK_LLM=0, the 2-step retry loop with automatic fallback prevents pipeline breaking in case of API throttling or schema validation errors.

    Part C — Multi-agent debate demo
        **Module Requirements & Scope**
        * Implement `debate.py` to run a 3-agent debate for one chosen stock from `STOCK_UNIVERSE` (e.g., `PAYTECH`).
        * **Required Agent Roles:**
            - Bull Agent: Argues for upside potential using the stock's return and beta.
            - Bear Agent: Argues against downside exposure using the stock's standard deviation (volatility).
            - Synthesizer Agent: Combines both viewpoints into a 2–3 sentence balanced summary.
        * Offline Mock Logic: Must run offline without external API calls and arguments must dynamically insert the selected stock's actual metrics (`beta`, `analyst_expected_return`, `std_dev`) into formatted text templates.
        * Run the script and print the full debate transcript.
        
        **Detailed Architecture Logic (debate.py)**
        The multi-agent debate system implemented in debate.py models an investment committee debate around a target asset (e.g., PAYTECH):
            1. Bull Agent: Focuses on upside potential, high systematic market participation (beta = 1.55), and projected return targets (E(R) = 19.0%). 
            2. Bear Agent: Focuses on tail risk, total risk exposure, and severe downside risk indicated by standard deviation (sigma = 34.0%). 
            3. Synthesizer Agent: Analyzes arguments from both sides and issues a 2–3 sentence risk-managed investment stance (e.g., HOLD/CAUTIOUS), matching the risk budget of the target investor profile. 
       
       Multi-Agent Debate Summary Table (PAYTECH):
            | Agent Role        | Input Quantitative Metrics     | Strategic Focus / Thesis                         |
            | Bull Agent        | Beta = 1.55, E(R) = 19.0%      | Growth, High Beta Momentum, Expected Upside      |
            | Bear Agent        | Std Dev = 34.0%                | Volatility, Downside Exposure, Macro Corrections |
            | Synthesizer Agent | Combined Matrix (Return, Risk) | Risk-Adjusted Allocation Stance                  |

        Execution Logs & Terminal Output
            =========================================================================================
                            MULTI-AGENT DEBATE DEMO: TICKER [PAYTECH]                     
            =========================================================================================
            Stock Metrics    : Beta = 1.55 | Expected Return = 19.0% | Volatility = 34.0%

            BULL AGENT: With an analyst expected return of 19.0% against a beta of 1.55, PAYTECH offers attractive upside potential and momentum to capture strong market rallies.
            -----------------------------------------------------------------------------------------
            BEAR AGENT: However, PAYTECH exhibits a high standard deviation of 34.0%, signaling significant volatility and downside exposure during broader market corrections.
            -----------------------------------------------------------------------------------------
            SYNTHESIZER: While PAYTECH provides strong return prospects (19.0%) amplified by a beta of 1.55, its elevated volatility (34.0%) requires strict risk controls. We synthesize a HOLD/CAUTIOUS stance, recommending exposure only for aggressive risk profiles.
            =========================================================================================

        **Implementation and Analysis**
            1. Bull Thesis (Upside Potential & Alpha Generation):
                * The bull agent leverages PAYTECH's high beta relative to the market benchmark. In an expanding economy or bull market, a beta of 1.55 acts as a return multiplier, allowing the asset to significantly outperform broad market indices.
                * The 19.0% expected return represents the highest projected yield across the stock universe, making a compelling growth case for return-maximizing capital.
            
            2. Bear Thesis (Unsystematic Volatility & Downside Vulnerability):
                * The bear agent isolates total risk, highlighting that a 34.0% annualized standard deviation creates severe drawdowns and downside tail risk during market shocks.
                * The high volatility dilutes the Sharpe ratio compared to lower-volatility assets in the portfolio (such as PAYGOLD or PAYRETAIL), indicating that the return per unit of total risk is relatively expensive.
            
            3. Synthesizer Resolution (Investment Committee Consensus):
                * The synthesizer correctly avoids a binary "Buy" or "Sell" recommendation, instead contextualizing the asset based on investor risk tolerance.
                * High-beta, high-volatility assets like PAYTECH are deemed suitable only for aggressive investor profiles (such as INV03 or INV05) while requiring strict risk controls (e.g., stop-loss limits or hard caps on portfolio weight) to prevent breach of risk budgets.    

    Part D — DCF valuation calculator

        **Module Requirements & Scope**
            * Implement `dcf_calculator.py` to calculate a Discounted Cash Flow (DCF) Enterprise Value (EV) for a hypothetical Paytm business line.
            * Calculate Base FCFF and Cost of Capital (WACC). 
            * Project FCFF for 5 years with a growth rate that gradually fades to a lower terminal growth rate.
            * Compute Terminal Value using the growing-perpetuity formula and discount all explicit cash flows and the terminal value back to present value at WACC to get Enterprise Value.
            * Set base terminal growth at least 3 percentage points below base WACC and verify the self-check condition: WACC − terminal_growth ≥ 1.
            * Generate a 3 x 3 grid varying WACC and terminal growth by 1.0 percentage point.
            * Compare the DCF Enterprise Value against an EV/EBITDA multiple estimation and provide a 2–3 sentence written comparison explaining the variance between the two valuations.

        **Detailed Architecture Logic (dcf_calculator.py)**
            The architectural design of the Discounted Cash Flow (DCF) Engine (`dcf_calculator.py`) is built as a deterministic, multi-stage valuation pipeline that calculates intrinsic enterprise value while enforcing strict mathematical boundary constraints. The pipeline begins by establishing baseline operating parameters, computing initial unlevered Free Cash Flow to the Firm (FCFF) using the standard formula FCFF = EBIT * (1 - tax rate) + Depreciation & Amortization - CapEx - Change in Net Working Capital. In parallel, the engine derives the Weighted Average Cost of Capital (WACC) by calculating the Cost of Equity via the Capital Asset Pricing Model (R_e = R_f + beta x (E(R_m) - R_f) using a target stock beta from `STOCK_UNIVERSE`, blending it with an after-tax cost of debt, and weighting both by an illustrative capital structure split (70% equity / 30% debt).

            In the projection and discounting phase, the engine models a 5-year explicit forecasting horizon where initial cash flow growth decelerates linearly toward a long-term terminal growth rate. The Year 6 normalized cash flow is then capitalized using the Gordon Growth perpetuity model TV = FCFF_6/(WACC - g_rate) to determine the Terminal Value. All explicit cash flows and the terminal value are discounted back to present value using the computed base WACC, yielding the baseline Intrinsic Enterprise Value (EV).

            To evaluate parameter sensitivity, the architecture generates a 3 x 3 sensitivity grid that varies WACC and terminal growth by 1.0 percentage point across both axes. A built-in mathematical self-check gate enforces that terminal growth is selected at least 3 percentage points below base WACC, guaranteeing that WACC - g_rate >= 1.0% even in the worst-case cell to prevent mathematical breakdown or division-by-zero errors. Finally, the engine cross-checks the intrinsic DCF result against a market-based valuation (EBITDA x Peer Multiple), automatically outputting a comparative analysis that contrasts public market multiples against long-term intrinsic cash flow fundamentals.

        **Parameters & Key Financial Inputs**
        All monetary figures are expressed in INR Crores (₹ Cr).
            - Base Unlevered Free Cash Flow (FCFF):
            FCFF = EBIT × (1 − tax rate) + D&A − CapEx − ΔNet Working Capital  
                 = 500.0 x (1 - 0.25) + 50.0 - 80.0 - 20.0 
                 = ₹325.00 Cr
        
            - Weighted Average Cost of Capital (WACC) Derivation:
                * Risk-Free Rate (R_f): 7.0% 
                * Expected Market Return (E(R_m)): 13.0% 
                * Equity Beta (beta_PAYFIN): 1.35
                * Cost of Equity (R_e): R_f + beta x (E(R_m) - R_f) 
                                      = 7.0% + 1.35 x (13.0% - 7.0%) 
                                      = 15.10%
                * Pre-Tax Cost of Debt (R_d): 8.0%
                * After-Tax Cost of Debt (R_d (after-tax)): 8.0% x (1 - 0.25) 
                                                          = 6.00%
                * Capital Structure Weights: 70% Equity / 30% Debt (w_e = 0.70, w_d = 0.30) 
                * Base WACC: (0.70 X 15.10%) + (0.30 X 6.00%) 
                           = 12.37% 
            - Growth Trajectory & Required Self-Check
                *  initial Growth Rate (g_init): 12.0% fading linearly to terminal rate across 5 years. 
                * Terminal Growth Rate (g_term): 4.0% 
                * Self-Check Verification: (Worst-Case Spread) = Min WACC - Max (g_term) 
                                                               = (12.37% - 1.00%) - (4.00% + 1.00%) 
                                                               = 11.37% - 5.00% 
                                                               = 6.37% (Constraint satisfied: $6.37% >= 1.00% requirement).  

            - 5-Year Explicit FCFF Projections & Discounting Schedule:
            Cash flows are discounted back to Present Value (PV) at the base WACC of 12.37%:
                | Period | Growth Rate ($g_t$) | Explicit FCFF (INR Cr) | Discount Factor | Present Value (PV) (INR Cr) |
                | Year 1 |       12.00%        |       ₹364.00          |     0.8900      |             ₹323.96         |
                | Year 2 |       10.00%        |       ₹400.40          |     0.7920      |             ₹317.12         |
                | Year 3 |        8.00%        |       ₹432.43          |     0.7048      |             ₹304.78         |
                | Year 4 |        6.00%        |       ₹458.38          |     0.6272      |             ₹287.40         |
                | Year 5 |        4.00%        |       ₹476.71          |     0.5582      |             ₹266.10         |
                | Sum of Explicit PV FCFFs**   |          —             |        —        |           ₹1,499.36 Cr      |

            - Terminal Value (TV) & Enterprise Value (EV) Calculation:
                * Normalized Year 6 FCFF: FCFF_5 X (1 + g_term) 
                                        = 476.71 X (1 + 0.04) 
                                        = ₹495.78 Cr
                                        
                * Terminal Value at Year 5: FCFF_6/ (WACC - g_term) 
                                          = 495.78 / (0.1237 - 0.04) 
                                          = ₹ 5,923.32 Cr
                * PV of Terminal Value: 5,923.30/ (1 + 0.1237)^5 
                                      = ₹ 3,306.08 Cr
                * Base Enterprise Value (EV): PV(FCFF_5) + PV(TV) 
                                            = 1,499.36 + 3,306.08
                                            = ₹ 4,805.44 Cr

            - 3 x 3 Enterprise Value Sensitivity Grid (INR Crores):
                | WACC \ Terminal Growth (g_term) |   g = 3.0%   | g = 4.0% (Base) |   g = 5.0%   |
                | WACC = 11.37%                   | ₹4869.18     | ₹5465.18        | ₹6248.03      |
                | WACC = 12.37%                   | ₹4342.71     | ₹4805.44        | ₹5393.49      |
                | WACC = 13.37%                   | ₹3917.94     | ₹4286.67        | ₹4743.30      |

            - Multiples Cross-Check & Valuation Comparison:
                1. Illustrative EBITDA: EBIT + D&A 
                                   = 500.0 + 50.0 
                                   = ₹ 550.00 Cr
                2. Peer Multiple: 12.0 X (EV/EBITDA)
                3. Multiple-Derived EV: 550.00 X 12.0 
                                   = ₹ 6,600.00 Cr
                4. DCF Base EV: ₹ 4,805.44 Cr

                Comparison Commentary
                The DCF base-case valuation of ₹4,805.44 Cr provides an intrinsic baseline, while the 12.0x EV/EBITDA peer multiple yields ₹6,600.00 Cr (a 37.3% market premium). This variance reflects public market multiples pricing in immediate scale and broader market momentum, whereas our DCF conservatively discounts future cash flows at a 
                12.37% WACC with a fading growth profile.   

        **Execution Output & Verification Logs**
            =========================================================================================
                                DISCOUNTED CASH FLOW (DCF) VALUATION ENGINE                         
            =========================================================================================
            Base FCFF (Year 0)        : ₹325.00 Cr
            Cost of Equity (PAYFIN β) : 15.10% (R_f=7.0%, E(R_m)=13.0%, β=1.35)
            After-Tax Cost of Debt    : 6.00% (Pre-Tax=8.0%, Tax=25.0%)
            Computed WACC (70/30 D/E) : 12.37%
            Terminal Growth Rate (g)  : 4.00%
            -----------------------------------------------------------------------------------------
            Explicit 5-Year FCFF Projections (INR Cr):
            Year 1: ₹364.00 Cr
            Year 2: ₹400.40 Cr
            Year 3: ₹432.43 Cr
            Year 4: ₹458.38 Cr
            Year 5: ₹476.71 Cr
            Sum of PV Explicit FCFFs  : ₹1499.36 Cr
            PV of Terminal Value (TV) : ₹3306.08 Cr
            BASE ENTERPRISE VALUE (EV): ₹4805.44 Cr
            -----------------------------------------------------------------------------------------

            3x3 ENTERPRISE VALUE SENSITIVITY GRID (INR Crores):
            WACC / g        | g = 3.0%        | g = 4.0%        | g = 5.0%       
            --------------------------------------------------------------------
            WACC = 11.37%  | ₹4869.18       | ₹5465.18       | ₹6248.03       |
            WACC = 12.37%  | ₹4342.71       | ₹4805.44       | ₹5393.49       |
            WACC = 13.37%  | ₹3917.94       | ₹4286.67       | ₹4743.30       |
            --------------------------------------------------------------------
            Required Self-Check Verification: Worst-Case Spread (Min WACC - Max g) = 6.37% (>= 1.00% Pass)
            -----------------------------------------------------------------------------------------

            =========================================================================================
                                EV/EBITDA MULTIPLE CROSS-CHECK & COMPARISON                         
            =========================================================================================
            Base EBITDA (EBIT + D&A)  : ₹550.00 Cr
            Selected Peer Multiple    : 12.0x
            Multiple-Based EV         : ₹6600.00 Cr
            DCF Base Case EV          : ₹4805.44 Cr
            -----------------------------------------------------------------------------------------
            Comparison Comment:
            The DCF base-case valuation of ₹4805.44 Cr provides an intrinsic baseline, while the 12.0x EV/EBITDA peer multiple yields ₹6600.00 Cr (a 37.3% market premium). This variance reflects public market multiples pricing in immediate scale and broader market momentum, whereas our DCF conservatively discounts future cash flows at a 12.37% WACC with a fading growth profile.
            =========================================================================================
        
        **Analytical Breakdown**
            - Valuation Highlights:
                * Base Enterprise Value (EV): The baseline intrinsic DCF valuation is ₹4,805.44 Cr.
                * Terminal Value Dominance: 
                    - PV of Terminal Value = ₹3,306.08 Cr (68.8% of Total EV)
                    - PV of Explicit 5-Year FCFFs = ₹1,499.36 Cr (31.2% of Total EV)
                This high concentration in terminal value is standard for growth-stage fintech business lines, highlighting that most of the economic value is generated as cash flows mature and normalize.

            - Sensitivity Analysis Highlights:
                * Base Case: ₹4,805.44 Cr at WACC = 12.37% and $g_term = 4.0%.
                * Bull Case (Low WACC / High Growth): ₹6,248.03 Cr at WACC = 11.37%, g = 5.0% .
                * Bear Case (High WACC / Low Growth): ₹3,917.94 Cr at WACC = 13.37%, g = 3.0% .
            WACC Sensitivity: Every 1.0 percentage point shift in WACC impacts Enterprise Value by approximately ₹400 Cr – ₹450 Cr (approx 8.5%-9.5%), proving that cost of capital changes (such as interest rate shifts or beta updates) drive significant swings in intrinsic value.

            - DCF vs. EV/EBITDA Multiple Variance
                |Valuation Methodology | Enterprise Value | Implied Variance vs. DCF | Primary Valuation Driver                 |
                |Intrinsic DCF (Base)  |   ₹4,805.44 Cr   |       Baseline (0.0%)    | Conservative - 4.0% g_term & 12.37% WACC |
                |EV/EBITDA Multiple    |   ₹6,600.00 Cr   |       +37.3% Premium     | Public market pricing of current revenue scale & market share |

            Public market multiples yield a 37.3% higher valuation (12.0x EBITDA) because they capture current investor sentiment, sector momentum, and market expectations for aggressive future margin expansion. In contrast, the DCF model provides a more conservative, risk-controlled intrinsic floor by explicitly projecting a 5-year growth deceleration path (from 12.0% down to 4.0%) and applying a robust 12.37% discount rate.

    Part E — Blockchain/crypto risk-analysis appendix (written, no code required)
        **Module Requirements & Scope**
        Draft blockchain_risk_note.md as a 600–900 word structured Markdown file.
           * Evaluate stablecoin risks (fiat-collateralized vs. algorithmic), tokenomics (dilution and supply schedules), and DAO governance vulnerabilities (whale concentration and smart contract exploits) before surfacing crypto data to retail users.
           * Recommend a justified 0% (or capped 1%–2%) crypto allocation for retail portfolios based on CAPM principles, lack of cash flows, heavy-tailed volatility, survivorship bias, and transaction costs.
           * Apply the T.A.N.G. framework to identify two key social-engineering risk vectors across UPI, lending, and wallet systems, pairing each with a real-time bank-side defense mechanism.

        **Detailed Architecture Logic (blockchain_risk_note.md)**
        The markdown document presents the formal risk-management appendix (blockchain_risk_note.md) for integrating digital assets, structuring advisory algorithms, and defending against fraud vectors within Paytm's integrated super-app ecosystem.

        **Interpretation and Analysis**
        - Key Takeaways (Interpretation):
            * Crypto Watchlist Controls: Paytm Money must enforce strict filters before displaying crypto assets to retail users. Algorithmic stablecoins (which risk systemic death-spirals) should be completely banned, while fiat-backed stablecoins must require verified 1:1 reserve audits. DeFi tokens require third-party smart contract audits and governance concentration disclosures to protect users from whale manipulation.

            * Asset Allocation Framework: Standard financial models (like CAPM) do not justify holding crypto because it lacks cash flows and intrinsic yield. Due to extreme volatility, heavy tail-risk, tax burdens (30% tax + 1% TDS), and friction costs, default portfolio allocation should be 0.0%. For aggressive investors, an opt-in cap of 1.0% to 2.0% is acceptable solely for low-correlation diversification.

            * Social Engineering Defenses (T.A.N.G. Framework):
             - Authority/Need Risk (Remote KYC Scams): Fraudsters use fake regulatory threats to trick users into screen-sharing and draining credit/UPI accounts. 
             Defense: Real-time API sensors detect overlay/screen-sharing tools and instantly freeze financial gateways.
             
             - Greed/Temptation Risk (Fake Yield/Arbitrage Schemes): Attackers lure users into transferring money to fraud syndicates. 
             Defense: Graph velocity algorithms track real-time transaction flows to automatically freeze mule accounts before funds are withdrawn.

        - Core Strategic Analysis:
            * The text outlines a defensive, "safety-first" roadmap for super-apps expanding into digital assets. It balances product innovation (giving retail users market visibility) with consumer protection by excluding high-risk, uncollateralized structures like algorithmic stablecoins.

            * Rather than completely rejecting digital assets, the methodology leverages Modern Portfolio Theory appropriately: treating crypto not as a core wealth-building asset, but as a tightly capped, asymmetric tail-risk option (1–2%) for high-risk profiles while protecting standard investors with a zero default weight.

            * Because Paytm combines credit lines, UPI, and wealth products in a single app, a security breach in one node can cascade into others. The proposed real-time defenses (screen-overlay sensors and graph velocity mule freezes) directly counter vector aggregation by stopping automated loan drawdowns and interbank capital flight before settlement.
