# ExcelProject
I created an Excel project to help the employees to track customer's orders more easily and efficiently.
```excel
I used data validation function here and the list option to enable selecting one company from the whole list.
```
![image](https://github.com/user-attachments/assets/e5da63d4-02ff-41a2-89d7-e21ed53de204)
```excel
I also used vlookup function. Thanks to that, when the employee chooses particular company, the contact, phone number, fax, address, city, region, postal code and country will show automatically.
```
![image](https://github.com/user-attachments/assets/5131df61-fd15-497e-8d6b-f31a3284a92a)
```excel
Moreover, I added IF function to the empty boxes. Due to that N/A message appears. (This is the situation when some data in database weren't populated in the first place).  
```
![image](https://github.com/user-attachments/assets/c890eae7-220d-4135-acab-456be5579e0c)
```excel
I also decided to use subtotal function to have the sum of orders, average order amount and the date of last order for chosen company.
```
![image](https://github.com/user-attachments/assets/69483a87-c2b5-4304-a298-7c8da59ead7e)
```excel
What's more, I created a small macro which will download above data automatically, after selecting particular company name, and give message ‘finished’.
```
![image](https://github.com/user-attachments/assets/a3460e6b-acc6-42bf-9b4b-152b6f6351ce)
```excel
Then, I created a pivot table and based on it, a diagram which would show the sum of order amount per year, for the company which was selected in the first place.
```
![image](https://github.com/user-attachments/assets/a3918e9e-bdbe-44a4-a5f1-dff6b8f2d3a1)
```excel
Furthermore, I wanted the diagram to reflect the data visible in the Order History table, after selecting particular company. Additionally, my aim was to make the diagram disappear, when the company, which was chosen, doesn’t have any data in the database. I used another macro, to let it connect and work automatically.
```
![image](https://github.com/user-attachments/assets/7ada2d1f-a76d-4015-9442-04fd6338b8c1)
![image](https://github.com/user-attachments/assets/354f4673-7d07-4b12-8103-4039ad31632b)
![image](https://github.com/user-attachments/assets/3ea2ae89-819d-48e5-be30-3563efc3c882)

```excel
Last but not least, I used slicer function, so after clicking particular month, the employee has the order amount per month for particular company.
```
![image](https://github.com/user-attachments/assets/65e9ef50-3cbd-43b3-a6b7-0907c97defe2)







