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
        self.setWindowTitle("Campaigner Upsell Sales Performance Dashboard")
        self.setGeometry(100, 50, 1500, 900)
        main_layout = QVBoxLayout()

        # TITLE
        title = QLabel("Campaigner Upsell Sales Performance Dashboard")
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

        self.load_button = QPushButton("Load Report")
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

        self.sales_card, self.kpi_sales = create_kpi_card("Total Sales (Booked)", "0")
        self.approved_card, self.kpi_approved = create_kpi_card("Approved Sales", "0")
        self.revenue_card, self.kpi_revenue = create_kpi_card("Approved Revenue", "₹0")
        self.enach_card, self.kpi_enach = create_kpi_card("ENACH %", "0%")
        self.form_card, self.kpi_form = create_kpi_card("Form Filling %", "0%")


        kpi_layout.addWidget(self.sales_card)
        kpi_layout.addWidget(self.approved_card)
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
            start = self.start_date.date().toString("yyyy-MM-dd") + " 00:00:00"
            end = self.end_date.date().toString("yyyy-MM-dd") + " 23:59:59"


            conn = get_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)

            query = """
            WITH latest_ins_status AS (
                SELECT insurance_id, status
                FROM (
                    SELECT
                        insurance_id,
                        status,
                        ROW_NUMBER() OVER (
                            PARTITION BY insurance_id
                            ORDER BY updated_at DESC
                        ) AS rn
                    FROM insurance_subscriptions
                    WHERE status IN (1,4,6,8)
                ) t
                WHERE rn = 1
            ),

            insurance_summary AS (
                SELECT
                    ub.insurance_id,
                    c.id AS campaign_id,

                    SUM(
                        CASE
                            WHEN ub.sign = -1 THEN -ub.amount
                            ELSE ub.amount
                        END
                    ) / 1.18 AS Revenue,

                    MAX(ub.status) AS latest_booking_status,

                    SUBSTRING_INDEX(
                        GROUP_CONCAT(
                            ub.created_by
                            ORDER BY ub.approved_at DESC, ub.created_at DESC
                        ),
                        ',', 1
                    ) AS agent_id

                FROM user_bookings ub
                JOIN campaigns c
                    ON c.id = ub.source_camp_id
                WHERE ub.name = 'Super Top-UP'
                  AND ub.updated_at BETWEEN %s AND %s
                GROUP BY ub.insurance_id, c.id
            )

            SELECT
                CONCAT(u.firstname, ' ', u.lastname) AS Agent,

              
                COUNT(*) AS Booked_Qty,

               
                SUM(
                    CASE
                        WHEN s.latest_booking_status = 1 THEN 1
                        ELSE 0
                    END
                ) AS Approved_Qty,

              
                SUM(
                    CASE
                        WHEN s.latest_booking_status = 1
                         AND lis.status = 1 THEN 1
                        ELSE 0
                    END
                ) AS Active_Enach,

                
                SUM(
                    CASE
                        WHEN s.latest_booking_status = 1 THEN 1
                        ELSE 0
                    END
                ) AS OTP_Eligible,

                SUM(
                    CASE
                        WHEN s.latest_booking_status = 1
                         AND i.is_verified IN (1,2) THEN 1
                        ELSE 0
                    END
                ) AS Form_Filled,

                
                SUM(
                    CASE
                        WHEN s.latest_booking_status = 1 THEN s.Revenue
                        ELSE 0
                    END
                ) AS Revenue

            FROM insurance_summary s
            JOIN insurances i
                ON i.id = s.insurance_id
            LEFT JOIN latest_ins_status lis
                ON lis.insurance_id = i.id
            LEFT JOIN users u
                ON u.id = s.agent_id

            GROUP BY Agent
            ORDER BY Booked_Qty DESC;

            """

            cursor.execute(query, (start, end))
            df = pd.DataFrame(cursor.fetchall())
            cursor.close()
            conn.close()

            if df.empty:
                QMessageBox.information(self, "No Data", "No records found.")
                return

            # ===== NUMERIC SAFETY =====
            for col in ["Booked_Qty","Approved_Qty","Active_Enach","OTP_Eligible","Form_Filled","Revenue"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            # ===== KPI =====
            total_booked = int(df["Booked_Qty"].sum())
            total_approved = int(df["Approved_Qty"].sum())
            total_revenue = int(df["Revenue"].sum())

            self.kpi_sales.setText(str(total_booked))
            self.kpi_approved.setText(str(total_approved))
            self.kpi_revenue.setText(f"₹{total_revenue:,}")

            self.kpi_enach.setText(
                f"{int((df['Active_Enach'].sum()/df['Approved_Qty'].sum())*100) if df['Approved_Qty'].sum() else 0}%"
            )
            self.kpi_form.setText(
                f"{int((df['Form_Filled'].sum()/df['OTP_Eligible'].sum())*100) if df['OTP_Eligible'].sum() else 0}%"
            )

            # ===== TL MAP =====
            tl_map = pd.read_csv(
                "https://docs.google.com/spreadsheets/d/e/2PACX-1vRpIFMOGraCOG3SshQ0JcAaS_LidGk7zO5T9_Gh4IY7dhqosOGNEmFPgJWTyWsBlPjcCr9S1fTmZmBk/pub?output=csv"
            )
            df = df.merge(tl_map, on="Agent", how="left")
            df["TL"] = df["TL"].fillna("Not Assigned")

            # ===== TL SUMMARY =====
            tl_df = df.groupby("TL", as_index=False).agg({
                "Booked_Qty":"sum",
                "Approved_Qty":"sum",
                "Active_Enach":"sum",
                "OTP_Eligible":"sum",
                "Form_Filled":"sum",
                "Revenue":"sum"
            })

            tl_df["ENACH_%"] = np.where(
                tl_df["Booked_Qty"]>0,
                (tl_df["Active_Enach"]/tl_df["Approved_Qty"]*100).round(0),
                0
            ).astype(int).astype(str)+"%"

            tl_df["Form_%"] = np.where(
                tl_df["OTP_Eligible"]>0,
                (tl_df["Form_Filled"]/tl_df["OTP_Eligible"]*100).round(0),
                0
            ).astype(int).astype(str)+"%"

            df["ENACH_%"] = tl_df["ENACH_%"].values[0] if not tl_df.empty else "0%"
            df["Form_%"] = tl_df["Form_%"].values[0] if not tl_df.empty else "0%"

            df["Revenue"] = df["Revenue"].round(0).astype(int).apply(lambda x:f"{x:,}")
            tl_df["Revenue"] = tl_df["Revenue"].round(0).astype(int).apply(lambda x:f"{x:,}")

            self.populate_table(self.tl_table, tl_df)
            self.populate_table(
                self.agent_table,
                df[["Agent","TL","Booked_Qty","Approved_Qty","Revenue","ENACH_%","Form_%"]]
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def populate_table(self, table, dataframe):
        table.setRowCount(len(dataframe))
        table.setColumnCount(len(dataframe.columns))
        table.setHorizontalHeaderLabels(dataframe.columns)
        for r in range(len(dataframe)):
            for c in range(len(dataframe.columns)):
                item = QTableWidgetItem(str(dataframe.iloc[r,c]))
                item.setTextAlignment(Qt.AlignCenter)
                table.setItem(r,c,item)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)


# ================= RUN =================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Dashboard()
    window.show()
    sys.exit(app.exec_())
