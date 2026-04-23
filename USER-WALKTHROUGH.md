# Logictics Centre User Walkthrough

This guide explains how people use Logictics Centre day to day. It is written for office staff, logistics staff, and drivers rather than for developers.

## What Logictics Centre is for

Logictics Centre helps the team:

- create pickup and delivery work in one shared live system
- assign work to drivers or leave it unassigned until dispatch is ready
- track stock received, issued, and requested for artwork
- keep drivers on their own route view while office users keep a full global view

## Sign In And Move Around

1. Open Logictics Centre in your browser.
2. Sign in with your name and password.
3. Use the tabs below the top bar to move between pages.
4. Watch for confirmation messages after you save, assign, email, or complete something.

The pages you see depend on your role.

| Role | Pages |
| --- | --- |
| Admin | Dashboard, Global List, Assignments, Stock, Network, Users, Driver Lists |
| Sales | Dashboard, Global List, Assignments, Stock, Driver Lists |
| Logistics | Dashboard, Stock |
| Driver | Route, Completed |

## Quick Daily Flow By Role

### Admin

1. Check `Dashboard` for the live overview.
2. Use `Users` to create or update accounts.
3. Use `Network` to maintain pickup, factory, and saved client delivery locations.
4. Use `Global List` to create entries and send or download the CSV.
5. Use `Assignments` to place unassigned work onto driver lists.
6. Use `Driver Lists` to review each driver's live route and recorded position.
7. Use `Stock` for stock control, QR labels, and artwork requests.

### Sales

1. Use `Global List` to create entries.
2. Use `Assignments` to assign queued work to drivers.
3. Use `Driver Lists` to review live driver queues.
4. Use `Global List` to download or email the CSV.
5. Use `Stock` as a read-only reference for incoming and on-hand stock.

### Logistics

1. Open `Stock`.
2. Review `Recent activity` for the latest arrivals and shipments.
3. Add new stock items when needed.
4. Log stock in and stock out movements.
5. Use QR labels or QR scanning to speed up stock handling.
6. Send artwork requests when new work is needed.

### Driver

1. Open `Route`.
2. Allow location access if you want the map to start from your current position.
3. Open each stop, navigate to it, and work through the assigned entries.
4. Mark items as `Picked up` and then complete them when delivered or dropped off.
5. Use `Flag issue` if something was not collected or not yet ready.
6. Use `Transfer` if another active driver should take over the item.
7. Review finished work on `Completed`.

## First-Time Admin Setup

If the database has no users yet, the app opens a `Create the first admin account` screen.

1. Enter your name.
2. Enter and confirm your password.
3. Select `Create admin`.

The first account created becomes the initial admin account.

## Common Concepts

### Entries

An entry is a piece of live work for the team. It can be a `Collection` or a `Delivery`.

### Unassigned Work

If you leave the driver as `Unassigned`, the entry stays in the queue until dispatch is ready to assign it.

### Priority Stops

Admins can mark an entry as a priority stop. Priority stops are highlighted across the app and are given special attention in route planning.

### Stock Description

When creating an entry, the `Stock description` should describe what the driver is moving. Put one stock item per line if you want Logictics Centre to create separate stock records.

## Admin Walkthrough

### Dashboard

The admin dashboard gives a live snapshot of:

- open entries
- unassigned entries
- drivers
- completed entries

Use the cards on this page as shortcuts into the main admin areas.

### Users

Use `Users` to create and manage accounts for:

- admin
- sales
- driver
- logistics

When creating or editing a user:

1. Enter the person's name.
2. Choose the role.
3. Set a password.
4. Add a driver phone number if the role is `Driver`.

From the user list, admins can:

- edit an account
- disable or enable an account
- delete an account

Use `Disable` when you want to remove someone from active use without deleting their history.

### Network

Use `Network` to maintain the pickup, factory, and saved client delivery locations used in the app.

Each location can include:

- location name
- location type: `Supplier`, `Factory`, `Both`, or `Client delivery`
- physical address
- optional latitude and longitude
- optional contact person
- optional contact number

Supplier, factory, and both locations feed pickup planning. Client delivery locations can be reused when creating delivery entries.

### Global List

Use `Global List` to create work and review the full live list.

To create a new entry:

1. Choose a driver or leave the item `Unassigned`.
2. Choose the pickup location.
3. Choose `Collection` or `Delivery`.
4. Choose the `Scheduled date` for when the item should appear on the driver list and daily export.
5. Choose a saved delivery location or add a one-off delivery address if it is a delivery.
6. Tick `Save this delivery address as a reusable client location` when the destination should be available for future deliveries.
7. Tick `Mark this as a priority stop` if needed.
8. Enter the quote number.
9. Add any sales order number, invoice number, or PO number if available.
10. Enter the stock description.
11. Add optional branding or notice text.
12. For collections, enable `Move collected stock to a factory` if required and select the destination factory.
13. If needed, tick `Admin override for duplicate or return stop`.
14. Select `Create entry`.

Important notes:

- saving an entry also creates matching stock items in the `Stock` page
- the daily CSV and CSV email only include entries scheduled for the live date shown in the app
- delivery entries require either a saved delivery location or a one-off delivery address
- the admin override can be used for duplicate quote entries or for sending a driver back to a stop already completed that day

The rest of the Global List page lets admins:

- search entries by location, reference, driver, stock description, and related details
- expand each pickup location to review grouped entries
- delete entries when required
- download the live CSV
- send a test email
- send a rollover test email
- email the live CSV
- clear all active priority markers
- clear all rollover markers after the day's carry-over report has been checked

### Assignments

Use `Assignments` to move active work onto driver lists.

On this page you can:

- filter by all drivers, unassigned items, or a specific driver
- review the current driver, pickup location, and creator
- change the driver for an active entry
- return an item to `Unassigned`
- use `Admin override` if the assignment needs to bypass duplicate or completed-stop protection

Unassigned entries appear first so the dispatch queue is easy to work through.

### Stock

Admins have full stock control.

Use `Stock` to:

- review recent arrivals and shipments from the last 24 hours
- add new stock items
- edit or delete stock items
- log stock in and stock out movements
- edit existing stock movements
- open, print, share, or copy QR labels
- scan QR labels to select the correct stock item quickly
- send artwork requests
- review the artwork request log

When adding a stock item, enter:

- stock description
- at least one reference: quote, sales order, invoice, or PO
- optional stock code
- opening stock quantity
- optional notes

Deleting a stock item also removes its movement and artwork history, so use delete carefully.

### Driver Lists

Use `Driver Lists` to review the live route pages for each driver.

This view shows:

- each driver's active entries
- the scheduled date shown on each entry
- the stop sequence grouped by pickup location
- priority stop markers
- duplicate order warnings
- estimated route distance

Admins also see a map of the last recorded driver positions when drivers have allowed location access on their route page.

## Sales Walkthrough

### Dashboard

The sales dashboard shows driver coverage, open work, and how many entries you have created.

### Global List

Sales users can create entries using the same main form as admins, with two important differences:

- sales cannot use the admin override
- sales cannot deliberately bypass duplicate checks or send a driver back to a stop already completed that day

Sales users can also:

- search the grouped global list
- edit active entries
- expand pickup locations to review entries
- download the CSV
- send a test email
- send a rollover test email
- email the CSV

Sales users cannot delete entries.

### Assignments

Sales can assign or reassign active work using the `Assignments` page.

Typical use:

1. Filter to `Unassigned` to work through the queue.
2. Choose a driver from the dropdown.
3. Select `Assign` or `Save`.

Sales cannot use the admin override.

### Stock

Sales has read-only stock access. This page is useful for checking:

- recent arrivals
- current on-hand stock
- movement history
- order references attached to stock items

Sales cannot change stock records.

### Driver Lists

Use `Driver Lists` to review the live queues per driver and confirm what is currently on each route.

Each entry now shows its scheduled date so future work can be assigned ahead of time without losing track of which day it belongs to.

## Logistics Walkthrough

### Dashboard

The logistics dashboard gives a quick stock summary:

- stock items
- on-hand quantity
- movement count
- artwork request count

### Stock

The stock page is the main logistics workspace.

Logistics users can:

- review recent stock activity
- add new stock items
- log stock in and stock out
- edit stock movements
- open QR labels for stock items
- scan QR labels to select items for movement logging
- send artwork requests
- review the artwork request history

Logistics users cannot edit or delete existing stock items. Those actions stay with admins.

#### Logging Stock In Or Out

1. Open `Log stock in or out`.
2. Select the stock item.
3. Choose `Stock in` or `Stock out`.
4. Enter the quantity.
5. If it is stock in, enter the supplier.
6. If it is stock out, choose the driver.
7. Add notes if needed.
8. Select `Save movement`.

Each movement is timestamped and linked to the person who recorded it.

#### Using QR Labels

From the stock register, open `QR` on an item to:

- print a label
- share the QR
- copy the QR value

From the movement panel, use `Scan QR` to select the stock item by:

- live camera scan
- QR image upload
- manual QR value entry

#### Sending Artwork Requests

1. Open `Email artwork request`.
2. Choose the stock item.
3. Enter the requested quantity.
4. Add notes such as size, finish, or due date.
5. Select `Send artwork request`.

If the email buttons are disabled, mail delivery has not been configured yet.

## Driver Walkthrough

### Route

The `Route` page only shows the entries assigned to your name.

This page includes:

- a route map
- the stop sequence grouped by pickup location
- priority stop markers
- the number of active stops
- estimated route distance

If you allow location access, the route can start from your current position. If not, the route starts from the Johannesburg Dispatch Hub.

### Working A Stop

1. Open a stop with `Show orders`.
2. Use `Navigate` to open Google Maps.
3. Review the order details, notices, and stock description.
4. Select `Picked up` once the item is on the vehicle.
5. Complete the item when the job is finished.

Completion labels depend on the entry type:

- deliveries complete as `Delivered to client`
- collections complete as `Dropped at office`
- collections marked for factory movement can also complete as `Dropped at factory`

Completed items leave the live route and move to the `Completed` page.

### Flagging A Problem

Use `Flag issue` if there is a problem at the stop.

The follow-up reasons available are:

- `Not collected`
- `Not yet ready`

Add a driver note so the office can see what happened.

### Transferring An Entry

Use `Transfer` if another active driver should take over the item.

1. Select `Transfer`.
2. Choose the new driver.
3. Select `Transfer item`.

When a driver transfers an item, the admin email inbox is notified.

### Completed

Use `Completed` to review the entries you have already finished. This is your record of work already done for the day or later follow-up.

## Search And Filtering Tips

### Global List Search

The Global List search can match:

- location name or address
- driver name
- quote, sales order, invoice, or PO references
- stock description
- notices and follow-up notes

### Stock Search

The stock search can match:

- quote number
- sales order number
- invoice number
- PO number
- stock code
- item name

### Network Search

The network search can match:

- location name
- type
- address
- contact person
- contact number

## Troubleshooting

### I cannot see the page I need

The most likely reason is your role. Logictics Centre only shows the pages allowed for your user type.

### The Email buttons are disabled

Email delivery is not configured on the server yet. You can still use the app, but email actions stay unavailable until mail setup is completed.

### I cannot choose a factory destination

The factory option only becomes available when the entry type is `Collection` and there is at least one location marked as `Factory` or `Both`.

### My route map is not starting from my position

Allow location access in your browser. If location is blocked, unavailable, or times out, the route starts from the Johannesburg Dispatch Hub instead.

### QR camera scan is not working

Use one of the fallback options:

- upload a QR image
- type the printed QR value manually

### A stock item is not editable for logistics

That is expected. Logistics can add new stock items and manage movements, but only admins can edit or delete existing stock items.

## Best Practices

- Leave work unassigned until dispatch is ready.
- Put each stock item on its own line in the entry form for cleaner stock records.
- Use notices and driver follow-up notes to keep handovers clear.
- Search before creating duplicates.
- Use `Completed` as the finished-work record instead of keeping items on the live route.
