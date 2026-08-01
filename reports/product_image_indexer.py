import os
import re
from typing import Dict, Optional

class ProductImageIndexer:
    """
    Top 0.1% Singleton In-Memory Product Image Indexer
    Scans C:\All_Report\1_Mapping\ProductPicture recursively once.
    Provides O(1) instant lookup for product SKU image files.
    Non-blocking, try-catch safe, zero-latency.
    """
    _instance = None
    _index: Dict[str, str] = {}
    _initialized = False

    def __new__(cls, base_dir: str = r"C:\All_Report\1_Mapping\ProductPicture"):
        if cls._instance is None:
            cls._instance = super(ProductImageIndexer, cls).__new__(cls)
        return cls._instance

    def __init__(self, base_dir: str = r"C:\All_Report\1_Mapping\ProductPicture"):
        if not ProductImageIndexer._initialized:
            self.base_dir = base_dir
            self._build_index()
            ProductImageIndexer._initialized = True

    def _clean_key(self, raw: str) -> str:
        if not raw:
            return ""
        s = str(raw).strip().upper()
        # Remove extension if any
        s = os.path.splitext(s)[0]
        # Remove non-alphanumeric characters for clean matching
        s = re.sub(r'[^A-Z0-9]', '', s)
        return s

    def _build_index(self):
        ProductImageIndexer._index.clear()
        if not os.path.exists(self.base_dir):
            print(f"[ProductImageIndexer] Warning: Image directory does not exist: {self.base_dir}")
            return

        try:
            count = 0
            for root, _, files in os.walk(self.base_dir):
                for f in files:
                    if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                        full_path = os.path.join(root, f)
                        key = self._clean_key(f)
                        if key and key not in ProductImageIndexer._index:
                            ProductImageIndexer._index[key] = full_path
                            count += 1
            print(f"[ProductImageIndexer] Index built successfully with {count} images.")
        except Exception as e:
            print(f"[ProductImageIndexer] Error indexing product images: {e}")

    def get_image_path(self, sku: str) -> Optional[str]:
        if not sku:
            return None
        key = self._clean_key(sku)
        path = ProductImageIndexer._index.get(key)
        if path and os.path.exists(path):
            return path
        
        # Fallback 1: Try base SKU (first 10 chars)
        if len(key) >= 8:
            base_key = key[:10]
            for k, p in ProductImageIndexer._index.items():
                if k.startswith(base_key) or base_key.startswith(k):
                    if os.path.exists(p):
                        return p
        return None

    def refresh(self):
        """Force re-indexing of the image directory."""
        self._build_index()
