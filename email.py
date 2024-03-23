import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QFileDialog, QListWidget, QLabel

class EmailApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Email Matrix Organizer")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.import_button = QPushButton("Import Email Data")
        self.import_button.clicked.connect(self.import_data)
        self.layout.addWidget(self.import_button)

        # Filter options
        self.filter_layout = QVBoxLayout()

        self.country_label = QLabel("Country:")
        self.country_list = QListWidget()
        self.country_list.setSelectionMode(QListWidget.MultiSelection)
        self.filter_layout.addWidget(self.country_label)
        self.filter_layout.addWidget(self.country_list)

        self.category_label = QLabel("Category:")
        self.category_list = QListWidget()
        self.category_list.setSelectionMode(QListWidget.MultiSelection)
        self.filter_layout.addWidget(self.category_label)
        self.filter_layout.addWidget(self.category_list)

        self.partner_label = QLabel("Partner:")
        self.partner_list = QListWidget()
        self.partner_list.setSelectionMode(QListWidget.MultiSelection)
        self.filter_layout.addWidget(self.partner_label)
        self.filter_layout.addWidget(self.partner_list)

        self.layout.addLayout(self.filter_layout)

        self.filter_button = QPushButton("Filter")
        self.filter_button.clicked.connect(self.filter_data)
        self.layout.addWidget(self.filter_button)

        self.email_table = QTableWidget()
        self.layout.addWidget(self.email_table)

        self.central_widget.setLayout(self.layout)

    def import_data(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel files (*.xlsx)")
        file_dialog.setViewMode(QFileDialog.Detail)
        file_path, _ = file_dialog.getOpenFileName(self, "Select Excel File", "", "Excel files (*.xlsx)")

        if file_path:
            self.df = pd.read_excel(file_path)

            # Populate filter options
            self.country_list.addItems(self.df['Country'].astype(str).unique())
            self.category_list.addItems(self.df['Category'].astype(str).unique())
            self.partner_list.addItems(self.df['Partner'].astype(str).unique())

            # Display all data in table initially
            self.display_data(self.df)

    def display_data(self, df):
        self.email_table.clear()
        self.email_table.setRowCount(df.shape[0])
        self.email_table.setColumnCount(df.shape[1])
        self.email_table.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                self.email_table.setItem(i, j, item)

    def filter_data(self):
        selected_countries = [item.text() for item in self.country_list.selectedItems()]
        selected_categories = [item.text() for item in self.category_list.selectedItems()]
        selected_partners = [item.text() for item in self.partner_list.selectedItems()]

        filtered_df = self.df[self.df['Country'].isin(selected_countries) & 
                              self.df['Category'].isin(selected_categories) & 
                              self.df['Partner'].isin(selected_partners)]
        self.display_data(filtered_df)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = EmailApp()
    window.show()
    sys.exit(app.exec_())
