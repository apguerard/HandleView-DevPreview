ej.base.enableRipple(true);

// Sidebar Initialization
var sidebarMenu = new ej.navigations.Sidebar({
  width: "220px",
  mediaQuery: "(min-width: 600px)",
  target: ".main-content",
  dockSize: "50px",
  enableDock: true,
});
sidebarMenu.appendTo("#sidebar-menu");
//end of Sidebar initialization
// Toggle the Sidebar
document.getElementById("hamburger").addEventListener("click", function () {
  sidebarMenu.toggle();
});

var mainMenuItems = [
  {
    text: "Cutomers & Orders",
    iconCss: "icon-picture icon",
    items: [
      { text: "Top Ten Orders by Sales Amount" },
      { text: "Customer Details" },
      { text: "Customer List" },
      { text: "Order Details" },
      { text: "Order List" },
    ],
  },
  {
    text: "Inventory & Purchasing",
    iconCss: "icon-bell-alt icon",
    items: [
      { text: "Inventory List" },
      { text: "Product Details" },
      { text: "Purchase Order Details" },
      { text: "Purchase Order List" },
    ],
  },
  {
    text: "Suppliers",
    iconCss: "icon-tag icon",
    items: [{ text: "Supplier Details" }, { text: "Supplier List" }],
  },
  {
    text: "Shippers",
    iconCss: "icon-globe icon",
    items: [{ text: "Shipper Details" }, { text: "Shipper List" }],
  },
  {
    text: "Reports",
    iconCss: "icon-bookmark icon",
    items: [
      { text: "Sales Reports Dialog" },
      { text: "Customer Address Book" },
      { text: "Customer Phone Book" },
      { text: "Employee Address Book" },
      { text: "Employee Phone Book" },
      { text: "Invoice" },
      { text: "Monthly Sales Report" },
      { text: "Product Category Sales by Month" },
      { text: "Product Sales by Category" },
      { text: "Product Sales by Total Revenue" },
      { text: "Product Sales Quantity by Employee" },
      { text: "Quarterly Sales Report" },
      { text: "Supplier Address Book" },
      { text: "Supplier Phone Book" },
      { text: "Top Ten Biggest Orders" },
      { text: "Yearly Sales Report" },
    ],
  },
  {
    text: "Employees",
    iconCss: "icon-user icon",
    items: [{ text: "Employee Details" }, { text: "Employee List" }],
  },
  {
    text: "Supporting Objects",
    iconCss: "icon-picture icon",
  },
];
var mainMenuObj = new ej.navigations.Menu(
  { items: mainMenuItems, orientation: "Vertical", cssClass: "dock-menu" },
  "#main-menubar"
);
var accountMenuItem = [
  {
    text: "Account",
    items: [{ text: "Profile" }, { text: "Sign out" }],
  },
];
// horizontal-menubar initialization
var horizontalMenuobj = new ej.navigations.Menu(
  { items: accountMenuItem, orientation: "Horizontal", cssClass: "dock-menu" },
  "#horizontal-menubar"
);
