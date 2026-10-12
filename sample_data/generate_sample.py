"""Generate a deterministic, slightly messy sales export (sample_data/sales.csv)."""
import random
from datetime import date, timedelta
from pathlib import Path

PRODUCTS = {"Laptop Stand": 39.9, "USB-C Hub": 54.5, "Wireless Mouse": 24.99, "Mechanical Keyboard": 89.0,
            "Webcam HD": 64.0, "Monitor 27in": 229.0, "Noise-Cancelling Headset": 129.5,
            "Desk Lamp": 34.25, "Laptop Bag": 49.0, "External SSD 1TB": 99.99, "Docking Station": 149.0,
            "Ergonomic Chair": 279.0}
REGIONS = ["North", "South", "East", "West"]


def main(rows: int = 500, seed: int = 42) -> None:
    rnd = random.Random(seed)
    lines = ["date,region,product,units,unit_price"]
    for _ in range(rows):
        d = date(2025, 1, 1) + timedelta(days=rnd.randrange(365))
        product = rnd.choice(list(PRODUCTS))
        price = round(PRODUCTS[product] * rnd.uniform(0.95, 1.05), 2)
        lines.append(f"{d},{rnd.choice(REGIONS)},{product},{rnd.randint(1, 20)},{price}")
    # realistic dirt: duplicates, whitespace, bad values
    lines += [lines[5], lines[17], lines[40]]
    lines += ["2025-03-14,  North ,Desk Lamp ,4,34.25",
              "2025-13-45,East,USB-C Hub,3,54.5",
              "2025-05-02,West,Webcam HD,-2,64.0",
              "2025-06-11,,Laptop Bag,5,49.0",
              "2025-07-21,South,Monitor 27in,abc,229.0",
              "2025-08-09,East,Desk Lamp,2,0",
              "2025-09-30,North,,1,10.0"]
    out = Path(__file__).with_name("sales.csv")
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {out} ({len(lines) - 1} rows)")


if __name__ == "__main__":
    main()
