import pandas as pd
import sqlite3

def load_excel_data_to_df(excel_file_path):
    # Load each sheet into a DataFrame with the correct sheet names
    df_policy = pd.read_excel(excel_file_path, sheet_name='POLICY', engine='openpyxl')
    df_claims = pd.read_excel(excel_file_path, sheet_name='CLAIMS', engine='openpyxl')
    df_extensions = pd.read_excel(excel_file_path, sheet_name='POL_SUB', engine='openpyxl')  # Assuming POL_SUB is your Extensions sheet
    df_agency = pd.read_excel(excel_file_path, sheet_name='AGENCY', engine='openpyxl')
    return df_policy, df_claims, df_extensions, df_agency


def transfer_df_to_sqlite(df_policy, df_claims, df_extensions, df_agency, conn):
    # Transfer the DataFrames into SQLite
    df_policy.to_sql('Policy', conn, index=False, if_exists='replace')
    df_claims.to_sql('Claims', conn, index=False, if_exists='replace')
    df_extensions.to_sql('Extensions', conn, index=False, if_exists='replace')
    df_agency.to_sql('Agency', conn, index=False, if_exists='replace')

# Updated version of the main function with corrected column names

def main():
    excel_file_path = 'data/datas.xlsx' 
    conn = sqlite3.connect(':memory:')  # Establish a connection to an in-memory SQLite database
    cursor = conn.cursor()  # Create a cursor object

    # Load Excel data into pandas DataFrames
    df_policy, df_claims, df_extensions, df_agency = load_excel_data_to_df(excel_file_path)

    # Transfer data from DataFrames into SQLite for SQL querying
    transfer_df_to_sqlite(df_policy, df_claims, df_extensions, df_agency, conn)

    # Example SQL for QA
    example_query = "SELECT * FROM Policy LIMIT 5;"  # Select first 5 rows from Policy table

    # Execute the example query
    cursor.execute(example_query)

    # Fetch and print the results of the example query
    rows = cursor.fetchall()
    print("Results of example QA query:")
    for row in rows:
        print(row)

    # Additional SQL queries
    queries = [

        # Premium income at quarter and year level
       """SELECT CAST(SUBSTR(MONTH_YEAR, 1, 4) AS INTEGER) AS Year,
                  CASE
                      WHEN CAST(SUBSTR(MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1'
                      WHEN CAST(SUBSTR(MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2'
                      WHEN CAST(SUBSTR(MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3'
                      ELSE 'Q4'
                  END AS Quarter,
                  SUM(TOTAL_PREM) AS PremiumIncome
           FROM Policy
           GROUP BY Year, Quarter
           ORDER BY Year, Quarter;""",

         # Annual premium income
    """SELECT CAST(SUBSTR(MONTH_YEAR, 1, 4) AS INTEGER) AS Year, 
              SUM(TOTAL_PREM) AS AnnualPremiumIncome 
       FROM Policy 
       GROUP BY Year;""",
 # Distribution of premium income by agency
    """SELECT AGENCY_ID, 
              SUM(TOTAL_PREM) AS PremiumIncome 
       FROM Policy 
       GROUP BY AGENCY_ID;""",

    # Total commissions paid to each top-down agency
    """SELECT AGENCY_ID, 
              SUM(SUB1 + SUB2 + SUB3 + SUB4 + SUB5) AS TotalCommissions 
       FROM Policy 
       GROUP BY AGENCY_ID;""",

    # Number of policies sold by each agency each year
    """SELECT CAST(SUBSTR(MONTH_YEAR, 1, 4) AS INTEGER) AS Year, 
              AGENCY_ID, 
              COUNT(*) AS PoliciesSold 
       FROM Policy 
       GROUP BY Year, AGENCY_ID;""",

    # The trend of the amount of expansions sold throughout all years.
  """SELECT CAST(SUBSTR(MONTH_YEAR, 1, 4) AS INTEGER) || '-' || SUBSTR(MONTH_YEAR, 5, 2) AS YearMonth, 
          SUM(SUB1 + SUB2 + SUB3 + SUB4 + SUB5) AS TotalExpansionsSold 
   FROM Policy 
   GROUP BY YearMonth;""",


# Corrected Query 7

"""SELECT 
  CAST(SUBSTR(Policy.MONTH_YEAR, 1, 4) AS INTEGER) || '-' || SUBSTR(Policy.MONTH_YEAR, 5, 2) AS YearMonth, 
  COUNT(*) AS NonProfitPolicies 
FROM 
  Policy 
INNER JOIN Claims ON Policy.POLICY_ID = Claims.POLICY_ID
WHERE 
  Policy.TOTAL_PREM < Claims.CLAIM_PAYMENT_NIS_AMOUNT 
GROUP BY 
  YearMonth;

""",


# another way for Query 8  - now assuming the column name for claim amounts is 'AMOUNT'
"""WITH Profitability AS ( 
    SELECT CAST(SUBSTR(Policy.MONTH_YEAR, 1, 4) AS INTEGER) AS Year, 
           CASE 
               WHEN CAST(SUBSTR(Policy.MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1' 
               WHEN CAST(SUBSTR(Policy.MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2' 
               WHEN CAST(SUBSTR(Policy.MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3' 
               ELSE 'Q4' 
           END AS Quarter, 
           SUM(Policy.TOTAL_PREM - Claims.CLAIM_PAYMENT_NIS_AMOUNT) AS TotalProfit, 
           (SUM(Policy.TOTAL_PREM - Claims.CLAIM_PAYMENT_NIS_AMOUNT) / SUM(Policy.TOTAL_PREM)) * 100 AS ProfitabilityRate 
    FROM Policy 
    JOIN Claims ON Policy.POLICY_ID = Claims.POLICY_ID
    GROUP BY Year, Quarter 
) 
SELECT Year, 
       Quarter, 
       TotalProfit, 
       ProfitabilityRate 
FROM Profitability;""",


# 8 by %
"""WITH Profitability AS ( 
    SELECT 
        CAST(SUBSTR(Policy.MONTH_YEAR, 1, 4) AS INTEGER) AS Year, 
        CASE 
            WHEN CAST(SUBSTR(Policy.MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1' 
            WHEN CAST(SUBSTR(Policy.MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2' 
            WHEN CAST(SUBSTR(Policy.MONTH_YEAR, 5, 2) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3' 
            ELSE 'Q4' 
        END AS Quarter, 
        SUM(Policy.TOTAL_PREM - Claims.CLAIM_PAYMENT_NIS_AMOUNT) AS TotalProfit,
        CASE
            WHEN SUM(Policy.TOTAL_PREM) > 0 THEN
                (SUM(Policy.TOTAL_PREM - Claims.CLAIM_PAYMENT_NIS_AMOUNT) / SUM(Policy.TOTAL_PREM)) * 100
            ELSE
                0
        END AS ProfitabilityRate 
    FROM Policy 
    INNER JOIN Claims ON Policy.POLICY_ID = Claims.POLICY_ID
    GROUP BY Year, Quarter 
) 
SELECT 
    Year, 
    Quarter, 
    TotalProfit, 
    ProfitabilityRate 
FROM Profitability;""",


# 10 total by year 
"""SELECT 
    CAST(SUBSTR(Policy.MONTH_YEAR, 1, 4) AS INTEGER) AS Year, 
    SUM(Policy.TOTAL_PREM - Claims.CLAIM_PAYMENT_NIS_AMOUNT) AS TotalProfitByYear,
    CASE
        WHEN SUM(Policy.TOTAL_PREM) > 0 THEN
            (SUM(Policy.TOTAL_PREM - Claims.CLAIM_PAYMENT_NIS_AMOUNT) / SUM(Policy.TOTAL_PREM)) * 100
        ELSE
            0
    END AS ProfitabilityRateByYear 
FROM Policy 
INNER JOIN Claims ON Policy.POLICY_ID = Claims.POLICY_ID
GROUP BY Year;""",

# Year-over-Year Customer Retention Rate per agancy
"""
WITH CustomerPurchases AS (
  SELECT
    AGENCY_ID,
    ID_NUM,
    CAST(SUBSTR(MONTH_YEAR, 1, 4) AS INTEGER) AS Year,
    COUNT(*) AS NumberOfPolicies
  FROM Policy
  GROUP BY AGENCY_ID, ID_NUM, Year
),
YearlyRetention AS (
  SELECT
    CP1.AGENCY_ID,
    CP1.Year AS PurchaseYear,
    COUNT(DISTINCT CP1.ID_NUM) AS TotalCustomers,
    COUNT(DISTINCT CP2.ID_NUM) AS RetainedCustomers,
    (COUNT(DISTINCT CP2.ID_NUM) * 100.0 / COUNT(DISTINCT CP1.ID_NUM)) AS RetentionRate -- Calculate here
  FROM CustomerPurchases CP1
  LEFT JOIN CustomerPurchases CP2 ON CP1.ID_NUM = CP2.ID_NUM AND CP1.AGENCY_ID = CP2.AGENCY_ID AND CP2.Year = CP1.Year + 1
  GROUP BY CP1.AGENCY_ID, CP1.Year
)
SELECT 
  AGENCY_ID, 
  PurchaseYear, 
  TotalCustomers, 
  RetainedCustomers, 
  RetentionRate
FROM YearlyRetention

UNION ALL
SELECT
  AGENCY_ID,
  'Total' AS PurchaseYear,
  SUM(TotalCustomers) AS TotalCustomers,
  SUM(RetainedCustomers) AS RetainedCustomers,
  AVG(RetentionRate) AS RetentionRate -- Use the calculated rate from above
FROM YearlyRetention
GROUP BY AGENCY_ID
ORDER BY AGENCY_ID, PurchaseYear;
""",

# Average Claim Amount per Claim by Year
"""
SELECT 
    CAST(SUBSTR(C.MONTH_YEAR, 1, 4) AS INTEGER) AS Year,
    AVG(C.CLAIM_PAYMENT_NIS_AMOUNT) AS AvgClaimAmount
FROM 
    Claims C
GROUP BY 
    Year
ORDER BY 
    Year;


""",
# Total Premiums Collected per Year
"""SELECT 
    CAST(SUBSTR(MONTH_YEAR, 1, 4) AS INTEGER) AS Year,
    SUM(TOTAL_PREM) AS TotalPremiumsCollected
FROM 
    Policy
GROUP BY 
    Year
ORDER BY 
    Year;
""",
]


    # Execute additional queries
    for i, query in enumerate(queries, start=1):
        cursor.execute(query)
        rows = cursor.fetchall()
        print(f"\nResults of query {i}:")
        for row in rows:
            print(row)


    # Don't forget to close the connection when done
    conn.close()

if __name__ == "__main__":
    main()
