# Practical-8-.net

# Inventory Management System Using ADO.NET & CRUD Operations

## 1. Project Overview

The Inventory Management System is a web-based application developed using ASP.NET MVC, C#, ADO.NET, and SQL Server.

The application allows users to manage product information and perform CRUD operations:

- **Create:** Add new products.
- **Read:** Display all products.
- **Update:** Modify existing product details.
- **Delete:** Remove products from the database.

## 2. Technologies Used

| Technology | Purpose |
|---|---|
| C# | Programming language |
| ASP.NET MVC | Web application framework |
| ADO.NET | Database connectivity |
| SQL Server | Database management |
| HTML | Web page structure |
| Razor | Dynamic MVC views |
| Visual Studio | Development environment |
| .NET Framework 4.8 | Application framework |

## 3. Project Structure

```text
InventoryManagementMVC/
│
├── Controllers/
│   └── ProductController.cs
│
├── Data/
│   └── ProductRepository.cs
│
├── Models/
│   └── Product.cs
│
├── Views/
│   └── Product/
│       ├── Index.cshtml
│       ├── Form.cshtml
│       └── Delete.cshtml
│
├── Content/
│   └── Site.css
│
├── Database/
│   └── Setup.sql
│
├── Web.config
│
├── Global.asax
│
└── README.md
```

## 4. Database Creation

**Database Name:** InventoryPractical8

**Table Name:** Products

### SQL Script

File: `Database/Setup.sql`

```sql
USE master;
GO

IF DB_ID(N'InventoryPractical8') IS NULL
    CREATE DATABASE InventoryPractical8;
GO

USE InventoryPractical8;
GO

IF OBJECT_ID(N'dbo.Products', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Products
    (
        Id INT IDENTITY(1,1) PRIMARY KEY,
        Name NVARCHAR(80) NOT NULL,
        Category NVARCHAR(50) NOT NULL,
        Quantity INT NOT NULL
            CHECK (Quantity BETWEEN 0 AND 1000000),
        Price DECIMAL(18,2) NOT NULL
            CHECK (Price BETWEEN 0.01 AND 9999999.99)
    );
END;
GO
```

### Database Structure

| Column | Data Type | Description |
|---|---|---|
| Id | INT | Primary Key, Auto Increment |
| Name | NVARCHAR(80) | Product Name |
| Category | NVARCHAR(50) | Product Category |
| Quantity | INT | Available Stock |
| Price | DECIMAL(18,2) | Product Price |

## 5. Product Model

File: `Models/Product.cs`

```csharp
using System.ComponentModel.DataAnnotations;

namespace InventoryManagement.Models
{
    public class Product
    {
        public int Id { get; set; }

        [Required, StringLength(80)]
        [Display(Name = "Product Name")]
        public string Name { get; set; }

        [Required, StringLength(50)]
        public string Category { get; set; }

        [Range(0, 1000000)]
        public int Quantity { get; set; }

        [Range(typeof(decimal), "0.01", "9999999.99")]
        public decimal Price { get; set; }
    }
}
```

## 6. ADO.NET Product Repository

File: `Data/ProductRepository.cs`

```csharp
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using InventoryManagement.Models;

namespace InventoryManagement.Data
{
    public class ProductRepository
    {
        private readonly string connectionString =
            ConfigurationManager.ConnectionStrings["InventoryDb"]
            .ConnectionString;

        // READ ALL PRODUCTS
        public List<Product> GetAll()
        {
            var products = new List<Product>();

            using (var connection =
                new SqlConnection(connectionString))
            using (var command = new SqlCommand(
                "SELECT Id, Name, Category, Quantity, Price " +
                "FROM Products ORDER BY Id", connection))
            {
                connection.Open();

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        products.Add(ReadProduct(reader));
                    }
                }
            }

            return products;
        }

        // GET PRODUCT BY ID
        public Product GetById(int id)
        {
            using (var connection =
                new SqlConnection(connectionString))
            using (var command = new SqlCommand(
                "SELECT Id, Name, Category, Quantity, Price " +
                "FROM Products WHERE Id=@Id", connection))
            {
                command.Parameters.Add("@Id", SqlDbType.Int).Value = id;

                connection.Open();

                using (var reader = command.ExecuteReader())
                {
                    return reader.Read()
                        ? ReadProduct(reader)
                        : null;
                }
            }
        }

        // ADD PRODUCT
        public void Add(Product product)
        {
            string query =
                "INSERT INTO Products " +
                "(Name, Category, Quantity, Price) " +
                "VALUES (@Name, @Category, @Quantity, @Price)";

            using (var connection =
                new SqlConnection(connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                AddProductParameters(command, product);

                connection.Open();
                command.ExecuteNonQuery();
            }
        }

        // UPDATE PRODUCT
        public bool Update(Product product)
        {
            string query =
                "UPDATE Products SET " +
                "Name=@Name, Category=@Category, " +
                "Quantity=@Quantity, Price=@Price " +
                "WHERE Id=@Id";

            using (var connection =
                new SqlConnection(connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.Add("@Id", SqlDbType.Int)
                    .Value = product.Id;

                AddProductParameters(command, product);

                connection.Open();

                return command.ExecuteNonQuery() > 0;
            }
        }

        // COMMON PRODUCT PARAMETERS
        private void AddProductParameters(
            SqlCommand command, Product product)
        {
            command.Parameters.Add(
                "@Name", SqlDbType.NVarChar, 80)
                .Value = product.Name;

            command.Parameters.Add(
                "@Category", SqlDbType.NVarChar, 50)
                .Value = product.Category;

            command.Parameters.Add(
                "@Quantity", SqlDbType.Int)
                .Value = product.Quantity;

            var price = command.Parameters.Add(
                "@Price", SqlDbType.Decimal);

            price.Precision = 18;
            price.Scale = 2;
            price.Value = product.Price;
        }

        // DELETE PRODUCT
        public bool Delete(int id)
        {
            using (var connection =
                new SqlConnection(connectionString))
            using (var command = new SqlCommand(
                "DELETE FROM Products WHERE Id=@Id", connection))
            {
                command.Parameters.Add("@Id", SqlDbType.Int)
                    .Value = id;

                connection.Open();

                return command.ExecuteNonQuery() > 0;
            }
        }

        // READ DATABASE RECORD
        private static Product ReadProduct(
            SqlDataReader reader)
        {
            return new Product
            {
                Id = reader.GetInt32(0),
                Name = reader.GetString(1),
                Category = reader.GetString(2),
                Quantity = reader.GetInt32(3),
                Price = reader.GetDecimal(4)
            };
        }
    }
}
```

## 7. Product Controller

File: `Controllers/ProductController.cs`

```csharp
using System.Web.Mvc;
using InventoryManagement.Data;
using InventoryManagement.Models;

namespace InventoryManagement.Controllers
{
    public class ProductController : Controller
    {
        private readonly ProductRepository repository =
            new ProductRepository();

        // DISPLAY ALL PRODUCTS
        public ActionResult Index()
        {
            return View(repository.GetAll());
        }

        // CREATE PRODUCT PAGE
        public ActionResult Create()
        {
            return View("Form", new Product());
        }

        // SAVE NEW PRODUCT
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(
            [Bind(Include = "Name,Category,Quantity,Price")]
            Product product)
        {
            if (!ModelState.IsValid)
                return View("Form", product);

            repository.Add(product);

            TempData["Success"] =
                "Product added successfully.";

            return RedirectToAction("Index");
        }

        // EDIT PRODUCT PAGE
        public ActionResult Edit(int id)
        {
            var product = repository.GetById(id);

            if (product == null)
                return HttpNotFound();

            return View("Form", product);
        }

        // UPDATE PRODUCT
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(
            int id,
            [Bind(Include = "Name,Category,Quantity,Price")]
            Product product)
        {
            product.Id = id;

            if (!ModelState.IsValid)
                return View("Form", product);

            if (!repository.Update(product))
                return HttpNotFound();

            TempData["Success"] =
                "Product updated successfully.";

            return RedirectToAction("Index");
        }

        // DELETE CONFIRMATION PAGE
        public ActionResult Delete(int id)
        {
            var product = repository.GetById(id);

            if (product == null)
                return HttpNotFound();

            return View(product);
        }

        // CONFIRM DELETE
        [HttpPost]
        [ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            if (!repository.Delete(id))
                return HttpNotFound();

            TempData["Success"] =
                "Product deleted successfully.";

            return RedirectToAction("Index");
        }
    }
}
```

## 8. Create Product View

File: `Views/Product/Form.cshtml`

```cshtml
@model InventoryManagement.Models.Product

@{
    var action = Model.Id == 0 ? "Create" : "Edit";
    ViewBag.Title = Model.Id == 0
        ? "Add Product"
        : "Edit Product";
}

<h1>@ViewBag.Title</h1>

<p>Enter product details. Price is in INR.</p>

@using (Html.BeginForm(
    action, "Product",
    new { id = Model.Id },
    FormMethod.Post,
    new { @class = "panel" }))
{
    @Html.AntiForgeryToken()

    @Html.ValidationSummary(
        false, "Please correct the following errors:")

    <div class="field">
        @Html.LabelFor(m => m.Name)
        @Html.TextBoxFor(m => m.Name,
            new { maxlength = 80 })
        @Html.ValidationMessageFor(m => m.Name)
    </div>

    <div class="field">
        @Html.LabelFor(m => m.Category)
        @Html.TextBoxFor(m => m.Category,
            new { maxlength = 50 })
        @Html.ValidationMessageFor(m => m.Category)
    </div>

    <div class="field">
        @Html.LabelFor(m => m.Quantity)
        @Html.TextBoxFor(m => m.Quantity)
        @Html.ValidationMessageFor(m => m.Quantity)
    </div>

    <div class="field">
        @Html.LabelFor(m => m.Price)
        @Html.TextBoxFor(m => m.Price, "{0:0.00}")
        @Html.ValidationMessageFor(m => m.Price)
    </div>

    <button type="submit">Save Product</button>

    @Html.ActionLink(
        "Cancel", "Index", null,
        new { @class = "secondary" })
}
```

## 9. Product List View

File: `Views/Product/Index.cshtml`

```cshtml
@model List<InventoryManagement.Models.Product>

@{
    ViewBag.Title = "Inventory";
}

<div class="heading">
    <div>
        <h1>Inventory Management</h1>
        <p>Manage product stock and prices.</p>
    </div>

    @Html.ActionLink(
        "+ Add Product", "Create", null,
        new { @class = "button" })
</div>

@if (TempData["Success"] != null)
{
    <div class="success">
        @TempData["Success"]
    </div>
}

<p>
    <strong>Total Products: @Model.Count</strong>
</p>

<div class="panel table-wrap">

    <table>
        <thead>
            <tr>
                <th>ID</th>
                <th>Product Name</th>
                <th>Category</th>
                <th>Quantity</th>
                <th>Price (INR)</th>
                <th>Actions</th>
            </tr>
        </thead>

        <tbody>
            @foreach (var product in Model)
            {
                <tr>
                    <td>@product.Id</td>
                    <td>@product.Name</td>
                    <td>@product.Category</td>
                    <td>@product.Quantity</td>
                    <td>@product.Price.ToString("F2")</td>

                    <td>
                        @Html.ActionLink(
                            "Edit", "Edit",
                            new { id = product.Id })

                        |

                        @Html.ActionLink(
                            "Delete", "Delete",
                            new { id = product.Id })
                    </td>
                </tr>
            }
        </tbody>
    </table>

    @if (Model.Count == 0)
    {
        <p>No products found. Add your first product.</p>
    }

</div>
```

## 10. Delete Product View

File: `Views/Product/Delete.cshtml`

```cshtml
@model InventoryManagement.Models.Product

@{
    ViewBag.Title = "Delete Product";
}

<h1>Delete Product</h1>

<div class="panel">

    <h2>@Model.Name</h2>

    <p>
        Category: @Model.Category |
        Quantity: @Model.Quantity
    </p>

    <p>
        Price: INR @Model.Price.ToString("F2")
    </p>

    <p>
        Are you sure you want to delete this product?
    </p>

    @using (Html.BeginForm(
        "Delete", "Product",
        new { id = Model.Id },
        FormMethod.Post))
    {
        @Html.AntiForgeryToken()

        <button type="submit" class="danger">
            Delete Product
        </button>

        @Html.ActionLink(
            "Cancel", "Index", null,
            new { @class = "secondary" })
    }

</div>
```

## 11. Database Connection Configuration

File: `Web.config`

Add the following connection string inside the existing `<configuration>` element:

```xml
<connectionStrings>
    <add name="InventoryDb"
         connectionString="Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=InventoryPractical8;Integrated Security=True;"
         providerName="System.Data.SqlClient" />
</connectionStrings>
```

Do not add another connectionStrings section if one already exists.

## 12. Application Execution

1. Open the project in Visual Studio.
2. Create the SQL Server database.
3. Create the Products table.
4. Add the Product Model.
5. Create the Product Repository.
6. Create the Product Controller.
7. Add the three MVC Views.
8. Configure the database connection string.
9. Build the solution.
10. Run the application using F5.


Replace PORT with the port assigned by Visual Studio.

## 13. CRUD Operations

| Operation | SQL Command | Description |
|---|---|---|
| Create | INSERT | Add new products |
| Read | SELECT | Display products |
| Update | UPDATE | Modify product details |
| Delete | DELETE | Remove products |

## 14. Expected Output

Sample product records:

| ID | Product Name | Category | Quantity | Price |
|---:|---|---|---:|---:|
| 1 | Laptop | Electronics | 10 | 45000.00 |
| 2 | Mouse | Accessories | 25 | 500.00 |
| 3 | Keyboard | Accessories | 15 | 1200.00 |

The application provides the following screens:

1. Add Product Form
2. Product List
3. Edit Product Form
4. Delete Confirmation
5. Updated Product List

## 15. Result

The Inventory Management Module was developed using ASP.NET MVC, C#, ADO.NET, and SQL Server.

The application implements Create, Read, Update, and Delete operations to manage product information efficiently.

## 16. Conclusion

This project demonstrates the use of ADO.NET for database connectivity in an ASP.NET MVC application. It provides practical knowledge of SQL Server, database operations, MVC architecture, and CRUD functionality.

---

**Project Name:** Inventory Management System

**Experiment:** 8

**Subject:** .NET Programming

**Framework:** .NET Framework 4.8

**Language:** C#

**Database:** SQL Server

**Database Access:** ADO.NET
