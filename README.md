> [!Note]
> This project aims to semi-automate the reconciliation of current product quantities from various data sources, i.e., multiple warehouses. Since the warehouses operate independently and are leased by the company, their systems are not interconnected. Warehouses are only responsible for receiving, storing, and shipping goods; they do not participate in product sales or the company's accounting system. However, it is essential to align the company's inventory records with the actual quantities reported by the warehouses.
> 
> Each month, the warehouses provide the company with a product inventory report. Due to their independence, the warehouses often use inconsistent naming conventions and classifications that differ from the company’s standards, making direct comparison impossible. Additionally, the company must compare and reconcile the quantities of hundreds of product styles across all warehouses. The goal of this project is to improve efficiency and accuracy in reconciling the company’s records with the actual quantities reported by the warehouses.

***Usage***

1. **Company Data Preparation**

    - the relevant data from QuickBooks that needs to be reconciled with warehouse records.

    - Process the data using [Durasein_cleaning.py](https://github.com/Uibelegia/GRL_reconciliation/blob/b5bcad530683c6a5e77f0d6498a240512da0c721/Durasein_cleaning.py) to generate the company’s inventory dataset for that specific warehouse.

2. **Warehouse Data Preparation**

    - Obtain the inventory reports from the respective warehouses.

    - As each warehouse report has its unique format, use the corresponding [{Warehouse_Location}_cleaning.py](https://pages.github.com/) script to clean and standardize the data.

3. **Reconciliation**

    - Save the cleaned datasets from both sources.

    - Use Reconciliation.py to generate a draft report that reconciles the company’s records with the warehouse’s actual inventory for hundreds of product styles.
