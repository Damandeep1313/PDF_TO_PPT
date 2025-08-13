import base64
import os
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from PIL import Image
from openai import OpenAI
from tqdm import tqdm

class ImageGenerator:
    def __init__(self, api_key, max_workers=10, cache_dir="img_cache"):
        self.client = OpenAI(api_key=api_key)
        self.max_workers = max_workers
        self.cache_dir = cache_dir
        os.makedirs(cache_dir, exist_ok=True)

    def _get_cache_path(self, prompt):
        """Generate consistent cache filename from prompt"""
        hash_obj = hashlib.sha256(prompt.encode())
        return os.path.join(self.cache_dir, f"{hash_obj.hexdigest()}.png")

    def _generate_single_image(self, prompt):
        """Core image generation logic"""
        cache_path = self._get_cache_path(prompt)
        
        # Return cached image if exists
        if os.path.exists(cache_path):
            return cache_path

        try:
            response = self.client.images.generate(
                model="gpt-image-1",
                prompt=f"{prompt}, professional corporate style",
                size="1024x1024"
            )

            # Handle response
            if hasattr(response.data[0], 'image'):  # Direct bytes
                img_data = response.data[0].image
            elif hasattr(response.data[0], 'b64_json'):  # Base64
                img_data = base64.b64decode(response.data[0].b64_json)
            else:
                raise ValueError("Unsupported image response format")

            # Save image
            with Image.open(BytesIO(img_data)) as img:
                img.save(cache_path)
            
            return cache_path

        except Exception as e:
            print(f"⚠️ Failed to generate image: {str(e)}")
            return None

    def generate_images(self, prompts):
        """
        Generate multiple images concurrently
        Args:
            prompts: List of prompt strings
        Returns:
            Dict of {prompt: image_path}
        """
        results = {}
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Create future->prompt mapping
            futures = {
                executor.submit(self._generate_single_image, p): p
                for p in prompts
            }
            
            # Process with progress bar
            for future in tqdm(
                as_completed(futures),
                total=len(prompts),
                desc="🎨 Generating images",
                unit="image"
            ):
                prompt = futures[future]
                results[prompt] = future.result()
        
        return results