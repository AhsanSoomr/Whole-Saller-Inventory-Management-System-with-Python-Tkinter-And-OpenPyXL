## Inventory Management System with Tkinter and OpenPyXL

### Overview
This project is a comprehensive Inventory Management System developed using Python's Tkinter library for the GUI and OpenPyXL for Excel file handling. It allows users to manage product inventory with functionalities to add, search, update, view, sell, and delete products. The product details, such as Product ID, Name, Description, Price, and Quantity, are stored in an Excel file, making data management easy and efficient.

### Features
- **Add Product**: Add new products to the inventory.
- **Search Product**: Search for products by name and display their details.
- **Update Product**: Update the details of existing products.
- **View Inventory**: View all products currently in the inventory.
- **Sell Product**: Sell products and update inventory quantities accordingly.
- **Delete Product**: Delete products from the inventory.
- **Clear Entries**: Clear input fields in the form.

### Requirements
- Python 3.x
- Tkinter (usually included with Python)
- OpenPyXL

### Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/AhsanSoomr/Whole-Saller-Inventory-Management-System-with-Python-Tkinter-And-OpenPyXL.git
   cd Whole-Saller-Inventory-Management-System-with-Python-Tkinter-And-OpenPyXL
   ```

2. Install the required packages:
   ```sh
   pip install openpyxl
   ```

### Usage
1. Run the application:
   ```sh
   python inventory_management.py
   ```

2. The main window will appear with fields to enter product details and buttons for various actions.

### Code Explanation
- **init_excel**: Initializes the Excel file with the necessary headers if it doesn't exist.
- **add_product**: Adds a new product to the inventory.
- **search_product**: Searches for a product by name and displays its details.
- **update_product**: Updates the details of an existing product.
- **view_inventory**: Displays all products in the inventory in a new window.
- **sell_product**: Sells a product and updates the quantity in the inventory.
- **delete_product**: Deletes a product from the inventory.
- **clear_entries**: Clears all input fields in the form.

### GUI Layout
- **Product ID**: Input field for the product ID.
- **Product Name**: Input field for the product name.
- **Product Description**: Input field for the product description.
- **Price**: Input field for the product price.
- **Quantity**: Input field for the product quantity.
- **Buttons**: Buttons for actions like Add, Search, Update, Clear, View Inventory, Sell, and Delete.



### Acknowledgements
- **Tkinter**: For providing the GUI components.
- **OpenPyXL**: For handling Excel file operations.

### Contact
For any inquiries or feedback, please contact [Your Name](mailto:ahsansoomro651@gmail.com).

### GitHub Repository
Check out the complete code and project details on GitHub: [Whole-Seller Inventory Management System with Python, Tkinter, and OpenPyXL](https://github.com/AhsanSoomr/Whole-Saller-Inventory-Management-System-with-Python-Tkinter-And-OpenPyXL).
