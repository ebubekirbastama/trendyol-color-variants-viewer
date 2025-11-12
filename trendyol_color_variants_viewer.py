"""
Trendyol Color Variants Viewer (Metro UI)
Author: Ebubekir Bastama
License: MIT

Description:
- Paste any Trendyol "color-variants" API URL and click "Fetch & Add".
- The app sends browser-like headers (User-Agent, Accept-Language, etc.) to avoid 400/403 errors.
- Results are displayed in a sortable data grid (ttk.Treeview).
- Live search by Product Name, ProductID, or barcode.
- Export the entire list to Excel (urunler.xlsx).
- Duplicate protection by ProductID.

Requirements:
    pip install customtkinter requests pandas openpyxl
"""

import threading
import json
import requests
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, messagebox

# ========================= App Settings =========================
APP_TITLE = "Trendyol Color Variants Viewer (Metro UI)"
DEFAULT_THEME = "dark"        # 'dark' | 'light' | 'system'
DEFAULT_COLOR = "green"       # 'blue' | 'green' | 'dark-blue'
WINDOW_SIZE = "1200x700"
MIN_SIZE = (1000, 600)
EXPORT_FILENAME = "urunler.xlsx"
# ===============================================================


def parse_color_variants(payload):
    """
    Flatten the expected structure:
    {
      "productGroupId": [ {item}, {item}, ... ],
      ...
    }
    Returns: list[dict] for DataFrame consumption.
    """
    rows = []
    if isinstance(payload, dict):
        for pg_id, items in payload.items():
            if not isinstance(items, list):
                continue
            for it in items:
                price = it.get("price", {}) or {}
                rating = it.get("ratingScore", {}) or {}
                rows.append({
                    "ProductGroupID": str(pg_id),
                    "ProductID": it.get("id"),
                    # barcode/mpn may not exist on this endpoint; keep for compatibility
                    "barcode": it.get("barcode") or it.get("mpn") or "",
                    "Product Name": it.get("name"),
                    "Price (TRY)": price.get("current"),
                    "Price Text": price.get("currentText"),
                    "Currency": price.get("currency"),
                    "Rating": rating.get("averageRating"),
                    "Review Count": rating.get("totalCount"),
                    "URL": ("https://www.trendyol.com" + it.get("url", "")) if it.get("url") else "",
                    "Image": it.get("image"),
                    "Big Image": it.get("bigImage"),
                    "Labels": ", ".join(it.get("labels", []) or []),
                })
    return rows


class TrendyolApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Metro (CustomTkinter) appearance
        ctk.set_appearance_mode(DEFAULT_THEME)
        ctk.set_default_color_theme(DEFAULT_COLOR)

        # Window
        self.title(APP_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(*MIN_SIZE)

        # In-memory table
        self.df = pd.DataFrame(columns=[
            "ProductGroupID", "ProductID", "barcode", "Product Name",
            "Price (TRY)", "Price Text", "Currency", "Rating", "Review Count",
            "URL", "Image", "Big Image", "Labels"
        ])
        self.seen_urls = set()   # optional: track processed URLs

        # ================= Top Bar (URL + Actions) =================
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.pack(side="top", fill="x", padx=10, pady=10)

        self.url_entry = ctk.CTkEntry(self.top_frame, placeholder_text="Paste a Trendyol color-variants URL...")
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)

        self.fetch_btn = ctk.CTkButton(self.top_frame, text="Fetch & Add", command=self.on_fetch_clicked)
        self.fetch_btn.pack(side="left", padx=5, pady=10)

        self.export_btn = ctk.CTkButton(self.top_frame, text="Export to Excel", command=self.on_export_clicked)
        self.export_btn.pack(side="left", padx=5, pady=10)

        self.clear_btn = ctk.CTkButton(
            self.top_frame, text="Clear", fg_color="#b91c1c",
            hover_color="#991b1b", command=self.on_clear_clicked
        )
        self.clear_btn.pack(side="left", padx=(5, 10), pady=10)

        # ================= Search Bar =================
        self.search_frame = ctk.CTkFrame(self)
        self.search_frame.pack(side="top", fill="x", padx=10)

        self.search_by_label = ctk.CTkLabel(self.search_frame, text="Search:")
        self.search_by_label.pack(side="left", padx=(10, 5), pady=(0, 8))

        self.search_by_var = ctk.StringVar(value="Product Name")
        self.search_by_combo = ctk.CTkComboBox(
            self.search_frame,
            values=["Product Name", "ProductID", "barcode"],
            variable=self.search_by_var,
            width=160
        )
        self.search_by_combo.pack(side="left", padx=5, pady=(0, 8))

        self.search_entry = ctk.CTkEntry(self.search_frame, placeholder_text="Type to filter...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5, pady=(0, 8))
        self.search_entry.bind("<KeyRelease>", self.on_search)

        self.count_label = ctk.CTkLabel(self.search_frame, text="Total: 0")
        self.count_label.pack(side="right", padx=10, pady=(0, 8))

        # ================= Data Grid =================
        self.grid_frame = ctk.CTkFrame(self)
        self.grid_frame.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 10))

        self.columns = [
            "ProductGroupID", "ProductID", "barcode", "Product Name",
            "Price (TRY)", "Rating", "Review Count", "Currency",
            "Price Text", "URL", "Image", "Big Image", "Labels"
        ]

        self.tree = ttk.Treeview(
            self.grid_frame, columns=self.columns, show="headings",
            selectmode="extended"
        )
        self.tree.pack(side="left", fill="both", expand=True)

        # Column headings and widths
        widths = {
            "ProductGroupID": 120, "ProductID": 120, "barcode": 140, "Product Name": 360,
            "Price (TRY)": 110, "Rating": 90, "Review Count": 120, "Currency": 100,
            "Price Text": 120, "URL": 340, "Image": 300, "Big Image": 300, "Labels": 180
        }
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c, False))
            self.tree.column(col, width=widths.get(col, 120), anchor="w")

        # Scrollbars
        self.v_scroll = ttk.Scrollbar(self.grid_frame, orient="vertical", command=self.tree.yview)
        self.v_scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.v_scroll.set)

        self.h_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.h_scroll.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=self.h_scroll.set)

        # ttk theme (clam for flat/metro alike)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", rowheight=28, borderwidth=0, relief="flat")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        # ================= Bottom (status + progress) =================
        self.bottom = ctk.CTkFrame(self)
        self.bottom.pack(side="bottom", fill="x", padx=10, pady=(0, 10))

        self.progress = ctk.CTkProgressBar(self.bottom)
        self.progress.set(0)
        self.progress.pack(side="left", fill="x", expand=True, padx=10, pady=8)

        self.status = ctk.CTkLabel(self.bottom, text="Ready.")
        self.status.pack(side="right", padx=10)

        # Initial counter
        self.update_counter()

    # ===================== Helpers =====================
    def set_status(self, text: str):
        self.status.configure(text=text)
        self.update_idletasks()

    def set_progress(self, value):
        try:
            self.progress.set(max(0.0, min(1.0, float(value))))
        except Exception:
            self.progress.set(0)

    def update_counter(self):
        self.count_label.configure(text=f"Total: {len(self.df)}")

    def refresh_tree(self, view_df: pd.DataFrame | None = None):
        """Refresh the grid with self.df or a filtered view_df."""
        target = view_df if view_df is not None else self.df
        # Clear
        for item in self.tree.get_children():
            self.tree.delete(item)
        # Refill
        for _, row in target.iterrows():
            vals = [row.get(c, "") for c in self.columns]
            self.tree.insert("", "end", values=vals)
        self.update_counter()

    # ===================== Events =====================
    def on_fetch_clicked(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Warning", "Please paste a Trendyol color-variants URL.")
            return

        # If you want to block re-processing the same URL, uncomment:
        # if url in self.seen_urls:
        #     messagebox.showinfo("Info", "This URL has already been processed.")
        #     return

        t = threading.Thread(target=self._fetch_thread, args=(url,), daemon=True)
        t.start()

    def _fetch_thread(self, url: str):
        """
        Fetch the given Trendyol color-variants endpoint.
        Uses realistic browser headers to avoid 400/403 responses.
        Optionally supply cookies if needed (e.g., AbTestingCookies).
        """
        try:
            self.set_status("Fetching...")
            self.set_progress(0.2)

            # Browser-like headers
            headers = {
                "accept": "application/json",
                "accept-language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
                "content-type": "application/json",
                "origin": "https://www.trendyol.com",
                "priority": "u=1, i",
                "referer": "https://www.trendyol.com/",
                "sec-ch-ua": '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-site",
                "user-agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/142.0.0.0 Safari/537.36"
                ),
                "x-request-source": "single-search-result",
            }

            # Minimal cookies usually work; paste your own from cURL if needed.
            cookies = {
                "platform": "web",
                "storefrontId": "1",
                "countryCode": "TR",
                "language": "tr",
            }

            resp = requests.get(url, headers=headers, cookies=cookies, timeout=40)
            resp.raise_for_status()

            self.set_progress(0.5)
            try:
                data = resp.json()
            except json.JSONDecodeError:
                data = json.loads(resp.text)

            rows = parse_color_variants(data)
            self.set_progress(0.7)

            if not rows:
                self.set_status("No data found or payload format is different.")
                self.set_progress(0)
                return

            new_df = pd.DataFrame(rows)

            # Merge & de-duplicate by ProductID
            combined = pd.concat([self.df, new_df], ignore_index=True)
            if "ProductID" in combined.columns:
                combined.drop_duplicates(subset=["ProductID"], keep="first", inplace=True)
            self.df = combined.reset_index(drop=True)

            self.refresh_tree()
            self.set_status("Done (200 OK).")
            self.set_progress(1.0)
            self.seen_urls.add(url)

        except requests.RequestException as e:
            self.set_status("Network error.")
            self.set_progress(0)
            messagebox.showerror("Error", f"Fetch error:\n{e}")
        except Exception as e:
            self.set_status("Error.")
            self.set_progress(0)
            messagebox.showerror("Error", f"Unexpected error:\n{e}")
        finally:
            # Smoothly reset progress
            self.after(800, lambda: self.set_progress(0))

    def on_export_clicked(self):
        if self.df.empty:
            messagebox.showinfo("Info", "There is no data to export.")
            return
        try:
            self.df.to_excel(EXPORT_FILENAME, index=False)
            messagebox.showinfo("Success", f"Excel saved: {EXPORT_FILENAME}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel:\n{e}")

    def on_clear_clicked(self):
        if messagebox.askyesno("Confirm", "Do you want to clear the list?"):
            self.df = self.df.iloc[0:0].copy()
            self.refresh_tree()
            self.seen_urls.clear()
            self.set_status("Cleared.")
            self.set_progress(0)

    def on_search(self, event=None):
        query = self.search_entry.get().strip().lower()
        field = self.search_by_var.get()

        if not query:
            self.refresh_tree(self.df)
            return

        field_map = {
            "Product Name": "Product Name",
            "ProductID": "ProductID",
            "barcode": "barcode",
        }
        col = field_map.get(field, "Product Name")
        if col not in self.df.columns:
            self.refresh_tree(self.df)
            return

        view_df = self.df[self.df[col].astype(str).str.lower().str.contains(query, na=False)]
        self.refresh_tree(view_df)

    # Column sorting
    def sort_by(self, col, descending):
        # Read current rows
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        # Convert numeric strings to float where possible for proper sorting
        def to_num(x):
            try:
                return float(str(x).replace(",", "."))
            except Exception:
                return x
        data.sort(key=lambda t: to_num(t[0]), reverse=descending)
        for index, (_, k) in enumerate(data):
            self.tree.move(k, "", index)
        # Toggle sort order on next click
        self.tree.heading(col, command=lambda _col=col: self.sort_by(_col, not descending))


if __name__ == "__main__":
    app = TrendyolApp()
    app.mainloop()
