#!/usr/bin/env python3
"""Merge order IDs from different sources"""

# Read existing orders
existing = set()
with open('lost-orders-woo.txt', 'r') as f:
    existing = {line.strip() for line in f if line.strip()}

# Read Gmail orders
new = set()
with open('gmail-orders.txt', 'r') as f:
    new = {line.strip() for line in f if line.strip()}

# Merge and deduplicate
combined = existing | new

# Sort numerically (longest/numbers first, then alphanumeric)
def sort_key(x):
    if x.isdigit():
        return (0, int(x))
    else:
        return (1, x)

sorted_ids = sorted(combined, key=sort_key, reverse=True)

# Write back
with open('lost-orders-woo.txt', 'w') as f:
    f.write('\n'.join(sorted_ids) + '\n')

print(f'Merged: {len(existing)} existing + {len(new)} new = {len(combined)} total unique IDs')
print(f'Saved to lost-orders-woo.txt')

