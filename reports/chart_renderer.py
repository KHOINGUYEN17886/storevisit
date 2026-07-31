import os
from PIL import Image, ImageDraw, ImageFont

class ChartRenderer:
    def __init__(self):
        pass

    def draw_revenue_chart(self, actual: int, target: int, prev: int, yoy: int, output_path: str):
        """
        Draw a premium, modern horizontal bar chart using Pillow.
        Saves the resulting image as PNG.
        """
        width, height = 1800, 500
        image = Image.new("RGBA", (width, height), (255, 255, 255, 0)) # transparent background
        draw = ImageDraw.Draw(image)

        # Harmonious color palette (HSL tailored / McKinsey style)
        colors = {
            "prev": (142, 172, 206, 255),      # Soft Blue
            "yoy": (180, 180, 180, 255),       # Grey
            "target": (237, 125, 49, 255),     # Warm Orange
            "actual": (10, 35, 66, 255)        # Deep Navy
        }

        labels = ["Thực tế lũy kế", "Kế hoạch tháng", "Cùng kỳ năm trước", "Tháng trước"]
        values = [actual, target, yoy, prev]
        
        # Calculate scale
        max_val = max(values) if max(values) > 0 else 1
        # Round up max_val for nice gridlines
        grid_max = max_val * 1.2
        
        # Geometry
        chart_left = 320  # Left space for labels
        chart_right = width - 200 # Right space for value labels
        chart_top = 50
        chart_bottom = height - 50
        chart_width = chart_right - chart_left
        chart_height = chart_bottom - chart_top

        # Draw grid lines (vertical)
        grid_lines = 4
        for i in range(grid_lines + 1):
            val = (grid_max / grid_lines) * i
            x = chart_left + (val / grid_max) * chart_width
            # Draw line
            draw.line([(x, chart_top), (x, chart_bottom)], fill=(220, 224, 230, 255), width=1)
            # Draw label
            label_text = f"{val / 1_000_000:.1f}M" if val >= 1_000_000 else f"{val / 1_000:.0f}K"
            if val == 0:
                label_text = "0"
            draw.text((x, chart_bottom + 15), label_text, fill=(120, 130, 140, 255), anchor="mt")

        # Draw horizontal bars
        num_bars = len(values)
        bar_gap = 35
        total_gaps_height = bar_gap * (num_bars + 1)
        bar_height = (chart_height - total_gaps_height) // num_bars

        for i in range(num_bars):
            val = values[i]
            col_key = ["actual", "target", "yoy", "prev"][i]
            col = colors[col_key]
            
            # Bar Y coordinates
            y_start = chart_top + bar_gap + i * (bar_height + bar_gap)
            y_end = y_start + bar_height
            
            # Bar X coordinates
            x_start = chart_left
            x_end = chart_left + (val / grid_max) * chart_width
            
            # Draw bar
            try:
                draw.rounded_rectangle([x_start, y_start, x_end, y_end], radius=6, fill=col)
            except AttributeError:
                draw.rectangle([x_start, y_start, x_end, y_end], fill=col)

            # Category labels on the left of the bar
            draw.text((chart_left - 20, (y_start + y_end) // 2), labels[i], fill=(40, 50, 60, 255), anchor="rm")
            
            # Draw values next to the bar
            val_str = f"{val:,.0f} đ"
            if col_key == "actual" and target > 0:
                pct = (actual / target) * 100
                val_str += f" ({pct:.1f}% Đạt)"
                
            draw.text((x_end + 15, (y_start + y_end) // 2), val_str, fill=(10, 35, 66, 255) if col_key == "actual" else (60, 70, 80, 255), anchor="lm")

        # Save image
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        image.save(output_path, "PNG")
        print(f"Rendered revenue horizontal chart to: {output_path}")


    def draw_stock_doughnut_chart(self, age_groups: dict, total_skus: int, output_path: str):
        """
        Draw a premium, modern McKinsey-style Doughnut Chart using Matplotlib.
        Saves the resulting image as PNG.
        """
        import matplotlib.pyplot as plt
        
        # Filter out zero values to keep chart clean
        plot_data = {k: v for k, v in age_groups.items() if v > 0}
        
        labels = list(plot_data.keys())
        sizes = list(plot_data.values())
        
        # McKinsey style color mapping
        color_map = {
            "Đợt PP tháng 7/2026": "#17a2b8",           # Teal
            "Quý 2/2026": "#4e73df",                    # Light Blue
            "Quý 1/2026": "#2e59d9",                    # Mid Blue
            "Quý 4/2025": "#f6c23e",                    # Soft Orange
            "Quý 3/2025": "#b7bec5",                    # Light Grey
            "Hàng nguyên giá PP > 1 năm": "#0a2342",     # Dark Navy
            "Hàng sale": "#e74a3b",                     # Crimson Red
            "Hàng thanh lý": "#85144b",                  # Deep Red
            "Khác/Chưa rõ": "#cccccc"
        }
        
        colors = [color_map.get(lbl, "#cccccc") for lbl in labels]
        
        # Create figure
        fig, ax = plt.subplots(figsize=(10, 8), subplot_kw=dict(aspect="equal"))
        
        # Draw pie chart
        wedges, texts, autotexts = ax.pie(
            sizes, 
            labels=labels, 
            autopct='%1.1f%%',
            startangle=90,
            colors=colors,
            textprops=dict(color="black", fontsize=12, fontweight='bold'),
            pctdistance=0.75,
            wedgeprops=dict(width=0.35, edgecolor='white', linewidth=2)  # width=0.35 creates doughnut hole
        )
        
        # Customize text labels
        for text in texts:
            text.set_fontsize(11)
            text.set_color("#2d3748")
        for autotext in autotexts:
            autotext.set_fontsize(10)
            autotext.set_color("white")
            
        # Draw total SKUs in the center hole
        ax.text(
            0, 0, 
            f"Tổng SKU\n{total_skus:,}", 
            ha='center', va='center', 
            fontsize=18, fontweight='bold', 
            color="#0a2342"
        )
        
        plt.tight_layout()
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        plt.savefig(output_path, dpi=200, transparent=True)
        plt.close()
        print(f"Rendered stock doughnut chart to: {output_path}")
