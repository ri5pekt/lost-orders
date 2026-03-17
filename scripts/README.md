# Legacy Scripts

Standalone Python scripts for batch-exporting Gmail invoice emails to PDF.

## Setup

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

## Export a single order

```bash
python export_one_order.py --order 3536793 --after 2025/11/01
```

Output: `out/order-3536793.html` + `out/order-3536793.pdf`

## Export many orders to one combined PDF

```bash
python export_orders_to_single_pdf.py --orders-file lost-orders-woo-3.txt --after 2025/11/01 --out-dir out-batch-3
```

Output: `out-batch-3/orders-combined.pdf`

## Extract order IDs from Gmail

```bash
python gmail_order_extractor.py
```
