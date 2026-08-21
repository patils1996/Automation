# BPCL Belgaum Territory Automation Daily Monitoring Dashboard

An elegant, fully responsive, and dynamic dashboard built for daily monitoring, tracking, and comparing the online uptime status of retail outlet technologies (Automation, ATG, EVCS, PINELAB/IOT). 

This application operates entirely in the browser (static frontend) and utilizes local **IndexedDB** for secure and persistent data logs. No backend server is required, making it 100% compatible with static hosting environments such as **GitHub Pages**.

## 🚀 Key Features

* **Neumorphic Interface**: Beautiful, responsive, soft-shadow neumorphic UI tailored for both **Laptops & Mobile screens**.
* **Yesterday vs Today Status Comparison**: A dedicated *Monitoring* tab comparing daily 10 PM status runs against today's active data to instantly spot drops (`Drop 🔴 ⬇️`) and upgrades (`Improvement 🟢 ⬆️`).
* **Dynamic Header Filters**: Instant filter dropdowns in both the *Day Monitoring* and *Monitoring* tables.
* **Role-Based Access Control (RBAC)**: Secure scoping for Admins, Territory Managers, Engineering Officers, Sales Officers, and Vendors (passwords and logs stored locally).
* **Tactile KPI Modals**: Clickable raised KPI buttons displaying detailed lists with custom CSV exporting.

---

## 🌐 Deploy to GitHub Pages (Access Anywhere / Mobile)

Since the app is purely frontend, you can deploy it to **GitHub Pages** in under a minute:

1. **Upload Code to GitHub**:
   - Create a new public repository on GitHub (e.g. `automation-dashboard`).
   - Push this directory's files to your repository:
     ```bash
     git add .
     git commit -m "Initial commit of responsive monitoring dashboard"
     git branch -M main
     git remote add origin https://github.com/YOUR_USERNAME/automation-dashboard.git
     git push -u origin main
     ```

2. **Enable GitHub Pages**:
   - Go to your repository on GitHub.com.
   - Click on the **Settings** tab.
   - Select **Pages** from the left-hand sidebar navigation.
   - Under *Build and deployment*, set the Source to **Deploy from a branch**.
   - Under *Branch*, select **main** (and `/root` folder), then click **Save**.

3. **Open the Dashboard**:
   - GitHub will generate a link for you, typically: `https://YOUR_USERNAME.github.io/automation-dashboard/`
   - You can immediately open this link on your **Mobile Phone, Tablet, or Laptop** to access the dashboard on the go!

---

## 🛠 Local Development / Host Locally
To run locally, simply start a static server in this directory:
```bash
python -m http.server 8000
```
Then open `http://localhost:8000` in your web browser.
