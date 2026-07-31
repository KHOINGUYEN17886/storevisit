import os
from PIL import Image, ImageOps, ImageDraw

class ImageProcessor:
    @staticmethod
    def process_and_fit_image(src_path: str, dest_path: str, target_width: int = 800, target_height: int = 600) -> str:
        """
        Processes an input image:
        1. Corrects orientation based on EXIF tags.
        2. Scales and fits the image inside target dimensions preserving aspect ratio (contain),
           centering it on a clean white background. No cropping or distortion.
        3. Saves as compressed JPEG.
        4. If input file is missing, creates a professional fallback placeholder.
        """
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        
        # Fallback if source image doesn't exist
        if not src_path or not os.path.exists(src_path):
            return ImageProcessor.create_placeholder(dest_path, "Hình ảnh chưa được cung cấp", target_width, target_height)

        try:
            with Image.open(src_path) as img:
                # 1. Correct orientation using EXIF data
                img = ImageOps.exif_transpose(img)
                
                # 2. Resize preserving aspect ratio (contain style)
                bg = Image.new("RGB", (target_width, target_height), (255, 255, 255))
                
                img_w, img_h = img.size
                ratio = min(target_width / img_w, target_height / img_h)
                new_w = int(img_w * ratio)
                new_h = int(img_h * ratio)
                
                img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                
                # Center on background
                offset_x = (target_width - new_w) // 2
                offset_y = (target_height - new_h) // 2
                bg.paste(img_resized, (offset_x, offset_y))
                
                # 3. Save as compressed JPEG
                bg.save(dest_path, "JPEG", quality=85)
                return dest_path
        except Exception as e:
            print(f"Warning: Failed to process image {src_path} due to: {e}. Using placeholder.")
            return ImageProcessor.create_placeholder(dest_path, f"Lỗi hình ảnh: {os.path.basename(src_path)}", target_width, target_height)

    @staticmethod
    def process_and_compress_image(src_path: str, dest_path: str, max_size: int = 1200) -> str:
        """
        Processes an image without cropping: orientates, scales to max_size, and compresses.
        """
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        if not src_path or not os.path.exists(src_path):
            return ImageProcessor.create_placeholder(dest_path, "Hình ảnh chưa được cung cấp", 800, 600)
            
        try:
            with Image.open(src_path) as img:
                img = ImageOps.exif_transpose(img)
                img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                img.convert("RGB").save(dest_path, "JPEG", quality=85)
                return dest_path
        except Exception as e:
            print(f"Warning: Failed to compress image {src_path} due to: {e}. Using placeholder.")
            return ImageProcessor.create_placeholder(dest_path, f"Lỗi hình ảnh: {os.path.basename(src_path)}", 800, 600)

    @staticmethod
    def create_placeholder(dest_path: str, text: str, width: int = 800, height: int = 600) -> str:
        """Create a professional placeholder image with text."""
        # Clean modern styling: Deep slate background with white text
        bg_color = (240, 243, 246)
        text_color = (120, 130, 140)
        border_color = (200, 208, 216)
        
        img = Image.new("RGB", (width, height), color=bg_color)
        draw = ImageDraw.Draw(img)
        
        # Draw soft border
        draw.rectangle([0, 0, width - 1, height - 1], outline=border_color, width=2)
        
        # Simple text draw
        # Calculate text position (center)
        # Using default font which is always available
        draw.text((width // 2, height // 2), text, fill=text_color, anchor="mm")
        
        img.save(dest_path, "JPEG", quality=80)
        return dest_path

    @staticmethod
    def create_multi_photo_grid(image_paths: list, dest_path: str, target_width: int = 1200, target_height: int = 900) -> str:
        """
        Combines 1 to 4 images into a clean, professional grid collage:
        - 1 image: Standard fit inside target bounds
        - 2 images: 2-column side-by-side split
        - 3-4 images: 2x2 grid collage with clean padding and subtle labels
        """
        valid_paths = [p for p in image_paths if p and os.path.exists(p) and os.path.getsize(p) > 0]
        if not valid_paths:
            return ImageProcessor.create_placeholder(dest_path, "Hình ảnh chưa được cung cấp", target_width, target_height)
            
        if len(valid_paths) == 1:
            return ImageProcessor.process_and_fit_image(valid_paths[0], dest_path, target_width, target_height)
            
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        canvas = Image.new("RGB", (target_width, target_height), (248, 250, 252))
        gap = 8
        
        if len(valid_paths) == 2:
            cell_w = (target_width - gap) // 2
            cell_h = target_height
            coords = [(0, 0), (cell_w + gap, 0)]
        else:
            cell_w = (target_width - gap) // 2
            cell_h = (target_height - gap) // 2
            coords = [
                (0, 0),
                (cell_w + gap, 0),
                (0, cell_h + gap),
                (cell_w + gap, cell_h + gap)
            ]
            
        for idx, src in enumerate(valid_paths[:4]):
            x, y = coords[idx]
            try:
                with Image.open(src) as img:
                    img = ImageOps.exif_transpose(img)
                    bg = Image.new("RGB", (cell_w, cell_h), (255, 255, 255))
                    img_w, img_h = img.size
                    ratio = min(cell_w / img_w, cell_h / img_h)
                    nw, nh = int(img_w * ratio), int(img_h * ratio)
                    img_resized = img.resize((nw, nh), Image.Resampling.LANCZOS)
                    ox = (cell_w - nw) // 2
                    oy = (cell_h - nh) // 2
                    bg.paste(img_resized, (ox, oy))
                    canvas.paste(bg, (x, y))
            except Exception as e:
                print(f"Warning: Grid photo {src} error: {e}")
                
        canvas.save(dest_path, "JPEG", quality=88)
        return dest_path
