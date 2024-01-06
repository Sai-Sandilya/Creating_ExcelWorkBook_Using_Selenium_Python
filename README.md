The openpyxl and selenium libraries are imported.
A new workbook is created using the openpyxl.Workbook() method.
A new worksheet is created using the workbook.active method and given the title “Data”.
The variables l and product_name are initialized.
Column headers are added to the worksheet using the ws.cell() method.
The webdriver module is used to open the Flipkart website and search for mobiles.
The product details such as product name, price, size, rating, image link, and product link are scraped from the website and saved to the worksheet using the ws.cell() method.
The workbook is saved using the workbook.save() method.
The browser window is closed using the driver.quit() method.
