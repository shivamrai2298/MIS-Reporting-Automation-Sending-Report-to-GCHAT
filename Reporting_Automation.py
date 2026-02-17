import sys
import pandas as pd
import numpy as np
import pymysql

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QDateEdit,
    QTableWidget, QTableWidgetItem,
    QHeaderView, QFrame, QMessageBox
)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QFont


# ================= DATABASE CONNECTION =================
def get_connection():
    return pymysql.connect(
        
)


# ================= DASHBOARD =================
class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IA Sales Performance Dashboard (Inbound - DIY)")
        self.setGeometry(100, 50, 1500, 900)
        main_layout = QVBoxLayout()

        # TITLE
        title = QLabel("IA SALES PERFORMANCE DASHBOARD")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        # DATE FILTER
        filter_layout = QHBoxLayout()
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addDays(-7))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())

        self.load_button = QPushButton("Load IA Data")
        self.load_button.setStyleSheet(
            "background-color: #2E86C1; color: white; padding: 8px; font-weight: bold;"
        )
        self.load_button.clicked.connect(self.load_data)

        filter_layout.addWidget(QLabel("Start Date"))
        filter_layout.addWidget(self.start_date)
        filter_layout.addWidget(QLabel("End Date"))
        filter_layout.addWidget(self.end_date)
        filter_layout.addWidget(self.load_button)
        main_layout.addLayout(filter_layout)

        # KPI CARDS
        kpi_layout = QHBoxLayout()
        def create_kpi_card(title, value):
            card = QFrame()
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border: 2px solid #D5D8DC;
                    border-radius: 12px;
                }
            """)
            card.setFixedHeight(120)
            layout = QVBoxLayout()
            title_label = QLabel(title)
            title_label.setFont(QFont("Arial", 11))
            title_label.setAlignment(Qt.AlignCenter)
            value_label = QLabel(value)
            value_label.setFont(QFont("Arial", 20, QFont.Bold))
            value_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(title_label)
            layout.addWidget(value_label)
            card.setLayout(layout)
            return card, value_label

        self.sales_card, self.kpi_sales = create_kpi_card("Total Sales", "0")
        self.revenue_card, self.kpi_revenue = create_kpi_card("Total Revenue", "₹0")
        self.enach_card, self.kpi_enach = create_kpi_card("ENACH %", "0%")
        self.form_card, self.kpi_form = create_kpi_card("Form Filling %", "0%")
        kpi_layout.addWidget(self.sales_card)
        kpi_layout.addWidget(self.revenue_card)
        kpi_layout.addWidget(self.enach_card)
        kpi_layout.addWidget(self.form_card)
        main_layout.addLayout(kpi_layout)

        # TL SUMMARY
        tl_label = QLabel("TL Wise Summary")
        tl_label.setFont(QFont("Arial", 14, QFont.Bold))
        main_layout.addWidget(tl_label)
        self.tl_table = QTableWidget()
        main_layout.addWidget(self.tl_table)

        # AGENT SUMMARY
        agent_label = QLabel("Agent Wise Performance")
        agent_label.setFont(QFont("Arial", 14, QFont.Bold))
        main_layout.addWidget(agent_label)
        self.agent_table = QTableWidget()
        main_layout.addWidget(self.agent_table)

        self.setLayout(main_layout)

    # ================= LOAD DATA =================
    def load_data(self):
        try:
            start = self.start_date.date().toString("yyyy-MM-dd")
            end = self.end_date.date().toString("yyyy-MM-dd")
            

            # ====== FETCH DATA FROM MYSQL ======
            try:
                conn = get_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)
                query = """
                WITH booking_combined AS (
                    SELECT source_camp_id,
                        SUM(CASE WHEN sign != -1 THEN amount ELSE 0 END)/1.18 AS total_amount,
                        SUM(CASE WHEN sign != -1 THEN 1 ELSE 0 END) AS booked_count,
                        MAX(CASE WHEN name = 'Campaigner Insurance Policy' THEN 1 ELSE 0 END) AS insurance_policy_present
                    FROM user_bookings
                    WHERE created_at BETWEEN %s AND %s
                      AND name IN ('Campaigner Policy', 'Campaigner Insurance Policy')
                    GROUP BY source_camp_id
                ),
                latest_ins_status AS (
                    SELECT insurance_id, status
                    FROM (
                        SELECT insurance_id, status,
                               ROW_NUMBER() OVER (PARTITION BY insurance_id ORDER BY updated_at DESC) AS rn
                        FROM insurance_subscriptions
                        WHERE status IN (1,4,6,8)
                    ) t
                    WHERE rn = 1
                ),
                base_data AS (
                    SELECT CONCAT(u.firstname,' ',u.lastname) AS Agent,
                           bc.total_amount AS Revenue,
                           CASE WHEN bc.booked_count > 0 THEN 1 ELSE 0 END AS Booked,
                           CASE WHEN lis.status = 1 THEN 1 ELSE 0 END AS Active_Enach,
                           CASE WHEN bc.insurance_policy_present = 1 THEN 1 ELSE 0 END AS OTP_Eligible,
                           CASE WHEN bc.insurance_policy_present = 1 AND i.is_verified IN (1,2) THEN 1 ELSE 0 END AS Form_Filled
                    FROM booking_combined bc
                    LEFT JOIN campaigns c ON c.id = bc.source_camp_id
                    LEFT JOIN campaign_extra_info cei ON cei.campaign_id = c.id
                    LEFT JOIN users u ON u.id = cei.pitched_by
                    LEFT JOIN insurances i ON i.module_id = c.id AND i.module_type = 3
                    LEFT JOIN latest_ins_status lis ON lis.insurance_id = i.id
                    WHERE c.is_marketing_campaign = 0
                      AND c.source_type BETWEEN 1 AND 6
                )
                SELECT Agent,
                       SUM(Revenue) AS Revenue,
                       SUM(Booked) AS Booked,
                       SUM(Active_Enach) AS Active_Enach,
                       SUM(OTP_Eligible) AS OTP_Eligible,
                       SUM(Form_Filled) AS Form_Filled
                FROM base_data
                GROUP BY Agent
                ORDER BY Booked DESC;
                """
                cursor.execute(query, (start, end))
                rows = cursor.fetchall()
                cursor.close()
                conn.close()
                
                    

            except Exception as e:
                rows = []
                
                QMessageBox.critical(self, "Database Error", f"MySQL fetch failed:\n{str(e)}")

            df = pd.DataFrame(rows)
            if df.empty:
                QMessageBox.information(self, "No Data", "No records found.")

            # ====== SAFE NUMERIC CONVERSION ======
            numeric_cols = ["Revenue", "Booked", "Active_Enach", "OTP_Eligible", "Form_Filled"]
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            # ====== PERCENTAGES ======
            for col, total_col, new_col in [("Booked", "Active_Enach", "ENACH_%"),
                                            ("OTP_Eligible", "Form_Filled", "Form_%")]:
                if col in df.columns and total_col in df.columns:
                    df[new_col] = np.where(df[col] > 0,
                                           (df[total_col] / df[col] * 100).round(0),
                                           0).astype(int).astype(str) + "%"
                else:
                    df[new_col] = []

            # ====== KPI CALCULATION ======
            total_sales = df["Booked"].sum() if "Booked" in df.columns else 0
            total_revenue = df["Revenue"].sum() if "Revenue" in df.columns else 0
            total_enach = df["Active_Enach"].sum() if "Active_Enach" in df.columns else 0
            total_eligible = df["OTP_Eligible"].sum() if "OTP_Eligible" in df.columns else 0
            total_form = df["Form_Filled"].sum() if "Form_Filled" in df.columns else 0

            self.kpi_sales.setText(str(int(total_sales)))
            self.kpi_revenue.setText(f"₹{int(total_revenue):,}")
            self.kpi_enach.setText(f"{int(round((total_enach/total_sales)*100 if total_sales else 0))}%")
            self.kpi_form.setText(f"{int(round((total_form/total_eligible)*100 if total_eligible else 0))}%")

            # ====== TL MAPPING ======
            sheet_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQMts0uUTsAl9TB3AIHyvrKm1PybUnNZYRrLkxXP0IGdmvIN0-8oZIYnr3xYrdto-o01P9L0WWN6Zjo/pub?gid=0&single=true&output=csv"
            try:
                tl_map = pd.read_csv(sheet_url)
            except Exception:
                tl_map = pd.DataFrame(columns=["Agent","TL"])

            

            if not df.empty and "Agent" in df.columns and "Agent" in tl_map.columns:
                df = df.merge(tl_map, on="Agent", how="left")
                df["TL"] = df["TL"].fillna("Not Assigned")
            else:
                df["TL"] = []

            # ====== TL SUMMARY ======
            tl_df = pd.DataFrame(columns=["TL","Booked","Revenue","Active_Enach","OTP_Eligible","Form_Filled"])
            if not df.empty and "TL" in df.columns:
                tl_df = df.groupby("TL", as_index=False).agg({
                    "Booked": "sum",
                    "Revenue": "sum",
                    "Active_Enach": "sum",
                    "OTP_Eligible": "sum",
                    "Form_Filled": "sum"
                })

            for col, total_col, new_col in [("Booked","Active_Enach","ENACH_%"),("OTP_Eligible","Form_Filled","Form_%")]:
                if col in tl_df.columns and total_col in tl_df.columns:
                    tl_df[new_col] = np.where(tl_df[col]>0,(tl_df[total_col]/tl_df[col]*100).round(0),0).astype(int).astype(str)+"%"
                else:
                    tl_df[new_col] = []

            for df_to_format in [df, tl_df]:
                if "Revenue" in df_to_format.columns and not df_to_format.empty:
                    df_to_format["Revenue"] = df_to_format["Revenue"].round(0).astype(int).apply(lambda x:f"{x:,}")

            # ====== POPULATE TABLES ======
            self.populate_table(self.tl_table, tl_df)
            agent_cols = ["Agent","TL","Booked","Revenue","ENACH_%","Form_%"]
            agent_cols = [col for col in agent_cols if col in df.columns]
            self.populate_table(self.agent_table, df[agent_cols])

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # ================= TABLE FUNCTION =================
    def populate_table(self, table, dataframe):
        table.setRowCount(len(dataframe))
        table.setColumnCount(len(dataframe.columns))
        table.setHorizontalHeaderLabels(dataframe.columns)
        for row in range(len(dataframe)):
            for col in range(len(dataframe.columns)):
                item = QTableWidgetItem(str(dataframe.iloc[row,col]))
                item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row,col,item)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setAlternatingRowColors(True)

# ================= RUN =================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Dashboard()
    window.show()
    sys.exit(app.exec_())

