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
        host="",
        user="",
        password="",
        database="",   
        port=3306
)


# ================= DASHBOARD =================
class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IA Sales Performance Dashboard (Outbound - DIY)")
        self.setGeometry(100, 50, 1500, 900)
        main_layout = QVBoxLayout()

        # TITLE
        title = QLabel("IA SALES PERFORMANCE DASHBOARD(Outbound - DIY)")
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

        self.sales_card,   self.kpi_sales   = create_kpi_card("Total Sales", "0")
        self.revenue_card, self.kpi_revenue = create_kpi_card("Total Revenue", "₹0")
        self.enach_card,   self.kpi_enach   = create_kpi_card("ENACH %", "0%")
        self.form_card,    self.kpi_form    = create_kpi_card("Form Filling %", "0%")

        kpi_layout.addWidget(self.sales_card)
        kpi_layout.addWidget(self.revenue_card)
        kpi_layout.addWidget(self.enach_card)
        kpi_layout.addWidget(self.form_card)
        main_layout.addLayout(kpi_layout)

        # CITY SUMMARY
        city_label = QLabel("City Wise Summary")
        city_label.setFont(QFont("Arial", 14, QFont.Bold))
        main_layout.addWidget(city_label)
        self.city_table = QTableWidget()
        main_layout.addWidget(self.city_table)

        # AGENT SUMMARY
        agent_label = QLabel("BDE Wise Performance")
        agent_label.setFont(QFont("Arial", 14, QFont.Bold))
        main_layout.addWidget(agent_label)
        self.agent_table = QTableWidget()
        main_layout.addWidget(self.agent_table)

        self.setLayout(main_layout)

    # ================= LOAD AGENT → CITY MAPPING =================
    def load_city_mapping(self):
        sheet_url = (
            "https://docs.google.com/spreadsheets/d/e/"
            "2PACX-1vQMts0uUTsAl9TB3AIHyvrKm1PybUnNZYRrLkxXP0IGdmvIN0-8oZIYnr3xYrdto-o01P9L0WWN6Zjo"
            "/pub?gid=0&single=true&output=csv"
        )
        try:
            raw = pd.read_csv(sheet_url, header=None)

            # Find the row where "Agent Name" and "City" headers appear
            header_row = None
            for i, row in raw.iterrows():
                val0 = str(row[0]).strip().lower()
                val1 = str(row[1]).strip().lower()
                if val0 == "agent name" and val1 == "city":
                    header_row = i
                    break

            if header_row is not None:
                city_map = raw.iloc[header_row + 1:, [0, 1]].copy()
                city_map.columns = ["Agent", "City"]
                city_map = city_map.dropna(subset=["Agent", "City"])
                city_map = city_map[city_map["Agent"].str.strip() != ""]
                city_map = city_map[city_map["City"].str.strip() != ""]
                city_map["Agent"] = city_map["Agent"].str.strip()
                city_map["City"]  = city_map["City"].str.strip()
                # Drop duplicates — keep first city per agent
                city_map = city_map.drop_duplicates(
                    subset=["Agent"], keep="first"
                ).reset_index(drop=True)
            else:
                city_map = pd.DataFrame(columns=["Agent", "City"])

        except Exception as e:
            QMessageBox.warning(
                self, "City Mapping Warning",
                f"Could not load city mapping sheet:\n{str(e)}\n\n"
                "All agents will show as 'Unknown City'."
            )
            city_map = pd.DataFrame(columns=["Agent", "City"])

        return city_map

    # ================= LOAD DATA =================
    def load_data(self):
        try:
            start = self.start_date.date().toString("yyyy-MM-dd")
            end   = self.end_date.date().toString("yyyy-MM-dd")

            # ====== FETCH DATA FROM MYSQL ======
            rows = []
            try:
                conn   = get_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)
                query  = """
                WITH booking_combined AS (
                    SELECT source_camp_id,
                        SUM(CASE WHEN sign != -1 THEN amount ELSE 0 END) / 1.18 AS total_amount,
                        SUM(CASE WHEN sign != -1 THEN 1 ELSE 0 END)              AS booked_count,
                        MAX(CASE WHEN name = 'Campaigner Insurance Policy' THEN 1 ELSE 0 END)
                                                                                  AS insurance_policy_present
                    FROM user_bookings
                    WHERE created_at BETWEEN %s AND %s
                      AND name IN ('Campaigner Policy', 'Campaigner Insurance Policy')
                    GROUP BY source_camp_id
                ),
                latest_ins_status AS (
                    SELECT insurance_id, status
                    FROM (
                        SELECT insurance_id, status,
                               ROW_NUMBER() OVER (
                                   PARTITION BY insurance_id ORDER BY updated_at DESC
                               ) AS rn
                        FROM insurance_subscriptions
                        WHERE status IN (1, 4, 6, 8)
                    ) t
                    WHERE rn = 1
                ),
                base_data AS (
                    SELECT CONCAT(u.firstname, ' ', u.lastname)                       AS Agent,
                           bc.total_amount                                             AS Revenue,
                           CASE WHEN bc.booked_count > 0 THEN 1 ELSE 0 END            AS Booked,
                           CASE WHEN lis.status = 1 THEN 1 ELSE 0 END                 AS Active_Enach,
                           CASE WHEN bc.insurance_policy_present = 1 THEN 1 ELSE 0 END
                                                                                       AS OTP_Eligible,
                           CASE WHEN bc.insurance_policy_present = 1
                                 AND i.is_verified IN (1, 2) THEN 1 ELSE 0 END        AS Form_Filled
                    FROM booking_combined bc
                    LEFT JOIN campaigns           c   ON c.id  = bc.source_camp_id
                    LEFT JOIN campaign_extra_info cei ON cei.campaign_id = c.id
                    LEFT JOIN users               u   ON u.id  = cei.pitched_by
                    LEFT JOIN insurances          i   ON i.module_id = c.id
                                                     AND i.module_type = 3
                    LEFT JOIN latest_ins_status   lis ON lis.insurance_id = i.id
                    WHERE c.is_marketing_campaign = 0
                      AND c.source_type = 0
                )
                SELECT Agent,
                       SUM(Revenue)      AS Revenue,
                       SUM(Booked)       AS Booked,
                       SUM(Active_Enach) AS Active_Enach,
                       SUM(OTP_Eligible) AS OTP_Eligible,
                       SUM(Form_Filled)  AS Form_Filled
                FROM base_data
                GROUP BY Agent
                ORDER BY Booked DESC;
                """
                cursor.execute(query, (start, end))
                rows = cursor.fetchall()
                cursor.close()
                conn.close()

            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"MySQL fetch failed:\n{str(e)}")

            df = pd.DataFrame(rows)

            if df.empty:
                QMessageBox.information(self, "No Data", "No records found for the selected date range.")
                return

            # ====== SAFE NUMERIC CONVERSION ======
            numeric_cols = ["Revenue", "Booked", "Active_Enach", "OTP_Eligible", "Form_Filled"]
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            # ====== CITY MAPPING ======
            city_map = self.load_city_mapping()

            if not df.empty and "Agent" in df.columns and not city_map.empty:
                # Case-insensitive merge via temporary lowercase key
                df["_agent_lower"]        = df["Agent"].str.strip().str.lower()
                city_map["_agent_lower"]  = city_map["Agent"].str.strip().str.lower()

                df = df.merge(
                    city_map[["_agent_lower", "City"]],
                    on="_agent_lower",
                    how="left"
                )
                df = df.drop(columns=["_agent_lower"])
                df["City"] = df["City"].fillna("Unknown City")
            else:
                df["City"] = "Unknown City"

            # ====== PERCENTAGES (agent level) ======
            pct_pairs = [
                ("Booked",       "Active_Enach", "ENACH_%"),
                ("OTP_Eligible", "Form_Filled",  "Form_%"),
            ]
            for denom_col, num_col, new_col in pct_pairs:
                if denom_col in df.columns and num_col in df.columns:
                    df[new_col] = np.where(
                        df[denom_col] > 0,
                        (df[num_col] / df[denom_col] * 100).round(0),
                        0
                    ).astype(int).astype(str) + "%"
                else:
                    df[new_col] = "0%"

            # ====== KPI CALCULATION ======
            total_sales    = df["Booked"].sum()       if "Booked"       in df.columns else 0
            total_revenue  = df["Revenue"].sum()      if "Revenue"      in df.columns else 0
            total_enach    = df["Active_Enach"].sum() if "Active_Enach" in df.columns else 0
            total_eligible = df["OTP_Eligible"].sum() if "OTP_Eligible" in df.columns else 0
            total_form     = df["Form_Filled"].sum()  if "Form_Filled"  in df.columns else 0

            self.kpi_sales.setText(str(int(total_sales)))
            self.kpi_revenue.setText(f"₹{int(total_revenue):,}")
            self.kpi_enach.setText(
                f"{int(round((total_enach / total_sales) * 100 if total_sales else 0))}%"
            )
            self.kpi_form.setText(
                f"{int(round((total_form / total_eligible) * 100 if total_eligible else 0))}%"
            )

            # ====== CITY SUMMARY ======
            city_df = pd.DataFrame()
            if "City" in df.columns:
                city_df = df.groupby("City", as_index=False).agg({
                    "Booked":       "sum",
                    "Revenue":      "sum",
                    "Active_Enach": "sum",
                    "OTP_Eligible": "sum",
                    "Form_Filled":  "sum",
                })
                city_df = city_df.sort_values(
                    "Booked", ascending=False
                ).reset_index(drop=True)

                # Percentages for City table
                for denom_col, num_col, new_col in pct_pairs:
                    if denom_col in city_df.columns and num_col in city_df.columns:
                        city_df[new_col] = np.where(
                            city_df[denom_col] > 0,
                            (city_df[num_col] / city_df[denom_col] * 100).round(0),
                            0
                        ).astype(int).astype(str) + "%"
                    else:
                        city_df[new_col] = "0%"

            # ====== FORMAT REVENUE ======
            for frame in [df, city_df]:
                if not frame.empty and "Revenue" in frame.columns:
                    frame["Revenue"] = (
                        frame["Revenue"]
                        .round(0)
                        .astype(int)
                        .apply(lambda x: f"₹{x:,}")
                    )

            # ====== POPULATE TABLES ======
            # City summary table
            city_cols = ["City", "Booked", "Revenue", "Active_Enach",
                         "OTP_Eligible", "ENACH_%", "Form_%"]
            city_cols = [c for c in city_cols if c in city_df.columns]
            self.populate_table(
                self.city_table,
                city_df[city_cols] if not city_df.empty
                else pd.DataFrame(columns=city_cols)
            )

            # Agent detail table
            agent_cols = ["Agent", "City", "Booked", "Revenue",
                          "Active_Enach", "OTP_Eligible", "ENACH_%", "Form_%"]
            agent_cols = [c for c in agent_cols if c in df.columns]
            self.populate_table(self.agent_table, df[agent_cols])

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Unexpected error:\n{str(e)}")

    # ================= POPULATE TABLE =================
    def populate_table(self, table, dataframe):
        if dataframe.empty:
            table.setRowCount(0)
            table.setColumnCount(0)
            return

        table.setRowCount(len(dataframe))
        table.setColumnCount(len(dataframe.columns))
        table.setHorizontalHeaderLabels(list(dataframe.columns))

        for row_idx in range(len(dataframe)):
            for col_idx in range(len(dataframe.columns)):
                value = str(dataframe.iloc[row_idx, col_idx])
                item  = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, col_idx, item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)


if __name__ == "__main__":
    app    = QApplication(sys.argv)
    window = Dashboard()
    window.show()
    sys.exit(app.exec_())
