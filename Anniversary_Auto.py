import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import win32com.client
from datetime import datetime
import os
import math
import requests
from io import BytesIO
import logging


class DynamicAnniversaryGenerator:
    def __init__(self, employee_data_file):
        """Initialize the Anniversary Generator with employee data"""
        # Load and clean the data
        self.employee_data = pd.read_excel(employee_data_file)
        
        # Clean column names (remove extra spaces)
        self.employee_data.columns = self.employee_data.columns.str.strip()
        
        # Validate data
        print("\n=== DATA VALIDATION ===")
        print(f"Total employees: {len(self.employee_data)}")
        print(f"Column names: {list(self.employee_data.columns)}")
        
        # Check for missing critical data
        critical_columns = ['Name', 'Department', 'JoiningDate', 'Email']
        for col in critical_columns:
            if col not in self.employee_data.columns:
                raise ValueError(f"Missing required column: {col}")
            
            missing_count = self.employee_data[col].isna().sum()
            if missing_count > 0:
                print(f"⚠️ Warning: {missing_count} employees have missing {col}")
        
        # Set canvas size and OFFICIAL company color scheme
        self.canvas_size = (1080, 1080)
        self.colors = {
            'brand_red': '#C0392B',        # Primary Brand Red
            'brand_red_alt': '#A93226',    # Alternative Brand Red
            'black': '#000000',          # Black for text
            'background': '#F5F5F5',     # Light gray background
            'white': '#FFFFFF',          # White
            'light_gray': '#E8E8E8',     # Light gray for boxes
            'medium_gray': '#CCCCCC'     # Medium gray for accents
        }
        
        # Company Logo path
        self.company_logo_path = r'C:\Projects\AnniversaryAutomation\CompanyLogo.png'
        
        # Setup logging
        self.setup_logging()
    
    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            filename='anniversary_automation.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
    
    def hex_to_rgb(self, hex_color):
        """Convert hex color to RGB tuple"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def get_dynamic_fonts(self, base_size=60):
        """Get fonts with fallback options - Times New Roman for title and name, Script for message"""
        # Times New Roman paths for title and name
        times_font_paths = [
            "C:/Windows/Fonts/timesbd.ttf",   # Times New Roman Bold
            "C:/Windows/Fonts/times.ttf",     # Times New Roman Regular
        ]
        
        # Script/Cursive font paths for congratulations message
        script_font_paths = [
            # ✨ MOST ELEGANT & CLASSY OPTIONS (Try these first!)
        "C:/Windows/Fonts/SCRIPTBL.TTF",   # Script MT Bold - Very elegant
        "C:/Windows/Fonts/FREESCPT.TTF",   # Freestyle Script - Flowing & beautiful
        "C:/Windows/Fonts/BRUSHSCI.TTF",   # Brush Script MT - Classic script
        "C:/Windows/Fonts/VLADIMIR.TTF",   # Vladimir Script - Very formal & classy
        "C:/Windows/Fonts/VIVALDII.TTF",   # Vivaldi - Extremely elegant calligraphy
        "C:/Windows/Fonts/KUNSTLER.TTF",   # Garamond (fallback)
        ]
        
        # Regular font paths for other text
        regular_font_paths = [
            "C:/Windows/Fonts/ariblk.ttf",   # Arial Black (bold)
            "C:/Windows/Fonts/arialbd.ttf",  # Arial Bold
            "C:/Windows/Fonts/calibrib.ttf", # Calibri Bold
            "C:/Windows/Fonts/arial.ttf",    # Arial Regular
        ]
        
        fonts = {}
        sizes = {
            'title': int(base_size * 1.1),      # Happy Anniversary (Times New Roman)
            'name': int(base_size * 1.0),       # Employee Name (Times New Roman)
            'department': int(base_size * 0.55), # Department
            'message': int(base_size * 0.50),    # Message (Script font - slightly larger)
            'years': int(base_size * 2.0),       # Large year number
        }
        
        # Load Times New Roman for title and name
        for font_type in ['title', 'name']:
            size = sizes[font_type]
            for font_path in times_font_paths:
                try:
                    fonts[font_type] = ImageFont.truetype(font_path, size)
                    print(f"✅ Loaded Times New Roman for {font_type}")
                    break
                except:
                    continue
            else:
                # Fallback to regular fonts if Times New Roman not found
                for font_path in regular_font_paths:
                    try:
                        fonts[font_type] = ImageFont.truetype(font_path, size)
                        print(f"⚠️ Times New Roman not found, using fallback for {font_type}")
                        break
                    except:
                        continue
                else:
                    fonts[font_type] = ImageFont.load_default()
        
        # Load Script font for message
        size = sizes['message']
        for font_path in script_font_paths:
            try:
                fonts['message'] = ImageFont.truetype(font_path, size)
                print(f"✅ Loaded Script font for message: {os.path.basename(font_path)}")
                break
            except:
                continue
        else:
            # Fallback to italic regular font if no script font found
            for font_path in regular_font_paths:
                try:
                    fonts['message'] = ImageFont.truetype(font_path, size)
                    print(f"⚠️ Script font not found, using fallback for message")
                    break
                except:
                    continue
            else:
                fonts['message'] = ImageFont.load_default()
        
        # Load regular fonts for other text
        for size_name in ['department', 'years']:
            size = sizes[size_name]
            for font_path in regular_font_paths:
                try:
                    fonts[size_name] = ImageFont.truetype(font_path, size)
                    break
                except:
                    continue
            else:
                fonts[size_name] = ImageFont.load_default()
        
        return fonts
    
    def load_company_logo(self, target_width=200):
        """Load and resize company logo"""
        try:
            # Try multiple possible extensions
            possible_paths = [
                self.company_logo_path,
                self.company_logo_path + '.png',
                self.company_logo_path + '.jpg',
                self.company_logo_path + '.jpeg',
                r'C:\Projects\AnniversaryAutomation\CompanyLogo.png',
                r'C:\Projects\AnniversaryAutomation\CompanyLogo.jpg',
            ]
            
            logo = None
            for path in possible_paths:
                if os.path.exists(path):
                    logo = Image.open(path)
                    print(f"✅ Company logo loaded from: {path}")
                    break
            
            if logo is None:
                print(f"⚠️ Company logo not found at any expected path")
                return None
            
            # Convert to RGBA if needed
            if logo.mode != 'RGBA':
                logo = logo.convert('RGBA')
            
            # Resize maintaining aspect ratio
            aspect_ratio = logo.height / logo.width
            target_height = int(target_width * aspect_ratio)
            logo = logo.resize((target_width, target_height), Image.Resampling.LANCZOS)
            
            print(f"✅ Logo resized to: {target_width}x{target_height}")
            return logo
            
        except Exception as e:
            print(f"❌ Error loading company logo: {e}")
            return None
        
    def create_background_pattern(self, canvas_size):
        """Create visible company circular symbol watermark pattern in lower curved area"""
        try:
            # Create transparent overlay
            pattern = Image.new('RGBA', canvas_size, (255, 255, 255, 0))
            
            # ===== LOAD COMPANY LOGO DIRECTLY FROM SPECIFIED PATH =====
            logo_path = r"C:\Projects\AnniversaryAutomation\photos\CompanyLogo_only.jpg"
            
            print(f"\n   🔍 Attempting to load company logo for pattern...")
            print(f"   📁 Path: {logo_path}")
            
            if not os.path.exists(logo_path):
                print(f"   ❌ Logo file not found at: {logo_path}")
                return pattern
            
            # Load the logo
            company_logo_full = Image.open(logo_path)
            print(f"   ✅ Logo loaded successfully!")
            
            # Convert to RGBA if needed
            if company_logo_full.mode != 'RGBA':
                company_logo_full = company_logo_full.convert('RGBA')
                print(f"   🔄 Converted from {Image.open(logo_path).mode} to RGBA")
            
            logo_width, logo_height = company_logo_full.size
            print(f"   📐 Original logo size: {logo_width}x{logo_height}px")
            
            # ===== EXTRACT ONLY THE CIRCULAR SYMBOL =====
            # Adjust this percentage based on your logo - the circular part is typically 40-50% of the width
            symbol_width = int(logo_width * 0.45)  # Adjust if needed
            
            # Crop to get the circular symbol (left portion of logo)
            company_symbol = company_logo_full.crop((0, 0, symbol_width, logo_height))
            print(f"   ✂️  Extracted circular symbol: {company_symbol.width}x{company_symbol.height}px")
            
            # Resize to watermark size
            watermark_size = 150  # Size of each symbol in pattern
            company_symbol = company_symbol.resize((watermark_size, watermark_size), Image.Resampling.LANCZOS)
            print(f"   📏 Resized symbol to: {watermark_size}x{watermark_size}px")
            
            # ===== MAKE IT VISIBLE BUT SUBTLE =====
            # Convert to grayscale
            symbol_gray = company_symbol.convert('L').convert('RGBA')
            
            # Apply watermark styling
            symbol_data = symbol_gray.getdata()
            watermark_data = []
            for item in symbol_data:
                if item[3] > 0:  # If pixel has content
                    # Light gray with visible opacity (you can adjust 100 to make more/less visible)
                    watermark_data.append((220, 220, 220, 100))  # Increased opacity to 100 for better visibility
                else:
                    watermark_data.append((255, 255, 255, 0))  # Transparent
            
            company_symbol.putdata(watermark_data)
            print(f"   🎨 Applied watermark styling (visible light gray)")
            
            # ===== DEFINE THE LOWER CURVED AREA =====
            width = canvas_size[0]
            height = canvas_size[1]
            
            # Lower curved area starts after the message text
            curve_start_y = int(height * 0.58)  # Start at 58% down
            bottom_clearance = 160  # Space for bottom company logo
            curve_end_y = height - bottom_clearance
            
            available_height = curve_end_y - curve_start_y
            
            print(f"   📍 Watermark area:")
            print(f"      Start Y: {curve_start_y}px")
            print(f"      End Y: {curve_end_y}px")
            print(f"      Available height: {available_height}px")
            
            # ===== CREATE SCATTERED PATTERN (12 symbols) =====
            positions = [
                # Left side
                (50, curve_start_y + 60),
                (30, curve_start_y + 200),
                (60, curve_end_y - 220),
                
                # Center-left
                (width // 4 - 50, curve_start_y + 120),
                (width // 4 - 30, curve_start_y + 280),
                (width // 4 - 60, curve_end_y - 180),
                
                # Center-right
                (width * 3 // 4 + 30, curve_start_y + 90),
                (width * 3 // 4 + 10, curve_start_y + 250),
                (width * 3 // 4 + 40, curve_end_y - 200),
                
                # Right side
                (width - watermark_size - 50, curve_start_y + 80),
                (width - watermark_size - 30, curve_start_y + 220),
                (width - watermark_size - 60, curve_end_y - 160),
            ]
            
            # ===== APPLY GRADIENT FADE =====
            center_x = width // 2
            center_y = (curve_start_y + curve_end_y) // 2
            
            symbols_placed = 0
            
            for idx, (x, y) in enumerate(positions):
                # Skip if outside defined area
                if y < curve_start_y or y > curve_end_y - watermark_size:
                    print(f"   ⚠️  Position {idx+1} outside area, skipping")
                    continue
                
                # Create copy for this position
                symbol_copy = company_symbol.copy()
                
                # Calculate distance from center for fade
                dx = (x + watermark_size/2 - center_x) / (width / 2)
                dy = (y + watermark_size/2 - center_y) / (available_height / 2)
                distance = (dx**2 + dy**2)**0.5 / (2**0.5)
                
                # Fade factor: 1.0 at center, 0.5 at edges
                fade_factor = max(0.5, 1.0 - (distance * 0.5))
                
                # Apply fade
                symbol_data_copy = symbol_copy.getdata()
                faded_data = []
                for item in symbol_data_copy:
                    if item[3] > 0:
                        new_opacity = int(item[3] * fade_factor)
                        faded_data.append((item[0], item[1], item[2], new_opacity))
                    else:
                        faded_data.append(item)
                
                symbol_copy.putdata(faded_data)
                
                # Paste onto pattern
                pattern.paste(symbol_copy, (x, y), symbol_copy)
                symbols_placed += 1
                print(f"   ✓ Symbol {symbols_placed} placed at ({x}, {y}), fade={fade_factor:.2f}")
            
            print(f"   ✅ Pattern complete: {symbols_placed} symbols placed in lower area")
            print(f"   ✅ Style: Circular symbol, visible watermark, edge-faded\n")
            
            return pattern
            
        except Exception as e:
            print(f"   ❌ Error creating watermark pattern: {e}")
            import traceback
            traceback.print_exc()
            return Image.new('RGBA', canvas_size, (255, 255, 255, 0))
        
    def get_ordinal_suffix(self, number):
        """Generate ordinal suffix for numbers (1st, 2nd, 3rd, etc.)"""
        if 10 <= number % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(number % 10, 'th')
        return f"{number}{suffix}"
    
    def calculate_anniversary_details(self, joining_date):
        """Calculate anniversary years and generate dynamic content"""
        today = datetime.now()
        years = today.year - joining_date.year
        if today.month < joining_date.month or (today.month == joining_date.month and today.day < joining_date.day):
            years -= 1
        
        suffix = self.get_ordinal_suffix(years).replace(str(years), '')
        
        # Dynamic messages based on years (shorter for better fit)
        messages = {
            1: "Congratulations on completing your first year with the company!\nWe're excited to see what's ahead!",
            2: "Two years of dedication and excellence!\nThank you for your valuable contributions!",
            3: "Three years of growth and success!\nYour impact continues to inspire us!",
            5: "Five wonderful years together!\nYour commitment makes our team stronger!",
            10: "A decade of outstanding service!\nYou're a true team champion!",
            15: "Fifteen years of excellence!\nYour legacy continues to grow!",
            20: "Twenty years of exceptional dedication!\nYou're a true legend!",
            25: "Twenty-five years of remarkable achievement!\nA true milestone to celebrate!"
        }
        
        if years in messages:
            message = messages[years]
        elif years < 5:
            message = f"Thank you for {years} {'year' if years == 1 else 'years'} of dedication!\nYour contributions make a real difference!"
        else:
            message = f"Celebrating {years} years of excellence!\nYour commitment to the team is truly valued!"
        
        return years, f"{years}{suffix}", message
    
    def create_placeholder_photo(self, target_size=(350, 450)):
        """Create a placeholder photo when employee photo is not available"""
        # Create a light gray background
        placeholder = Image.new('RGB', target_size, self.hex_to_rgb(self.colors['light_gray']))
        draw = ImageDraw.Draw(placeholder)
        
        # Draw a circle in the center with brand red
        circle_color = self.hex_to_rgb(self.colors['brand_red'])
        center_x, center_y = target_size[0] // 2, target_size[1] // 2
        radius = min(target_size) // 3
        
        draw.ellipse(
            [center_x - radius, center_y - radius, center_x + radius, center_y + radius],
            fill=circle_color
        )
        
        # Draw simple user icon (head and shoulders in white)
        head_radius = radius // 3
        draw.ellipse(
            [center_x - head_radius, center_y - radius//2 - head_radius, 
             center_x + head_radius, center_y - radius//2 + head_radius],
            fill='white'
        )
        
        # Add white border
        border_size = 10
        border_photo = Image.new('RGB', (target_size[0] + border_size, target_size[1] + border_size), 'white')
        border_photo.paste(placeholder, (border_size//2, border_size//2))
        
        return border_photo
    
    def process_employee_photo(self, photo_path, target_size=(350, 450)):
        """Process and prepare employee photo with enhanced debugging"""
        try:
            # Validate inpuCongratulationst
            if pd.isna(photo_path) or str(photo_path).strip() == '':
                print(f"   ❌ Photo path is empty or None")
                raise ValueError("Photo path is empty")
            
            photo_path = str(photo_path).strip()
            print(f"\n   📸 Original photo path from Excel: '{photo_path}'")
            
            # Check if it's a URL
            if photo_path.startswith('http://') or photo_path.startswith('https://'):
                print(f"   🌐 Detected URL, attempting download...")
                try:
                    response = requests.get(photo_path, timeout=10)
                    response.raise_for_status()
                    photo = Image.open(BytesIO(response.content))
                    print(f"   ✅ Photo downloaded successfully from URL")
                except Exception as url_error:
                    print(f"   ❌ Failed to download from URL: {url_error}")
                    raise
            else:
                # Handle local file path
                print(f"   📁 Detected local path, trying multiple locations...")
                
                # Get directories
                script_dir = os.path.dirname(os.path.abspath(__file__))
                excel_dir = r'C:\Projects\AnniversaryAutomation'
                
                # Create list of possible paths
                possible_paths = [
                    photo_path,
                    photo_path.replace('/', '\\'),
                    photo_path.replace('\\', '/'),
                    os.path.abspath(photo_path),
                    os.path.join(script_dir, photo_path),
                    os.path.join(script_dir, os.path.basename(photo_path)),
                    os.path.join(excel_dir, photo_path),
                    os.path.join(excel_dir, os.path.basename(photo_path)),
                    os.path.join(excel_dir, 'photos', os.path.basename(photo_path)),
                    os.path.join(excel_dir, 'images', os.path.basename(photo_path)),
                    os.path.join(excel_dir, 'Pictures', os.path.basename(photo_path)),
                ]
                
                # Add extension variations
                if '.' not in os.path.basename(photo_path):
                    possible_paths.extend([
                        photo_path + '.jpg',
                        photo_path + '.png',
                        photo_path + '.jpeg'
                    ])
                
                # Remove duplicates
                possible_paths = list(dict.fromkeys(possible_paths))
                
                print(f"   🔍 Attempting {len(possible_paths)} different path combinations...")
                
                photo = None
                working_path = None
                
                for idx, path_attempt in enumerate(possible_paths, 1):
                    normalized_path = os.path.normpath(path_attempt)
                    print(f"   [{idx}/{len(possible_paths)}] Trying: {normalized_path}")
                    
                    if os.path.exists(normalized_path):
                        try:
                            photo = Image.open(normalized_path)
                            working_path = normalized_path
                            print(f"   ✅ SUCCESS! Photo found and opened at: {normalized_path}")
                            break
                        except Exception as open_error:
                            print(f"   ⚠️  File exists but couldn't open: {open_error}")
                    else:
                        print(f"   ❌ File not found")
                
                if photo is None:
                    print(f"\n   ❌ PHOTO NOT FOUND after trying {len(possible_paths)} paths")
                    raise FileNotFoundError(f"Photo not found after trying {len(possible_paths)} different paths")
            
            # Convert to RGB if needed
            if photo.mode != 'RGB':
                print(f"   🔄 Converting image from {photo.mode} to RGB")
                photo = photo.convert('RGB')
            
            original_size = photo.size
            print(f"   📏 Original photo size: {original_size[0]}x{original_size[1]}")
            
            # Crop to target aspect ratio
            width, height = photo.size
            target_ratio = target_size[0] / target_size[1]
            current_ratio = width / height
            
            if current_ratio > target_ratio:
                new_width = int(height * target_ratio)
                left = (width - new_width) // 2
                photo = photo.crop((left, 0, left + new_width, height))
                print(f"   ✂️  Cropped width from {width} to {new_width}")
            elif current_ratio < target_ratio:
                new_height = int(width / target_ratio)
                top = (height - new_height) // 2
                photo = photo.crop((0, top, width, top + new_height))
                print(f"   ✂️  Cropped height from {height} to {new_height}")
            
            # Resize to target size
            photo = photo.resize(target_size, Image.Resampling.LANCZOS)
            print(f"   📐 Resized to: {target_size[0]}x{target_size[1]}")
            
            # Add white border
            border_size = 10
            border_photo = Image.new('RGB', (target_size[0] + border_size, target_size[1] + border_size), 'white')
            border_photo.paste(photo, (border_size//2, border_size//2))
            print(f"   🖼️  Added {border_size}px white border")
            
            print(f"   ✅ Photo processed successfully!")
            return border_photo
        
        except Exception as e:
            print(f"\n   ❌ ERROR processing photo: {type(e).__name__}")
            print(f"   📝 Error details: {str(e)}")
            print(f"   🎨 Creating placeholder image instead...")
            return self.create_placeholder_photo(target_size)
    
    def draw_centered_text(self, draw, text, y_position, font, color, canvas_width):
        """Helper function to draw centered text"""
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        x = (canvas_width - text_width) // 2
        draw.text((x, y_position), text, fill=color, font=font)
        return x, text_width
    
    def wrap_text(self, text, font, max_width, draw):
        """Wrap text to fit within max_width"""
        words = text.split()
        lines = []
        current_line = []
        
        for word in words:
            test_line = ' '.join(current_line + [word])
            bbox = draw.textbbox((0, 0), test_line, font=font)
            width = bbox[2] - bbox[0]
            
            if width <= max_width:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                current_line = [word]
        
        if current_line:
            lines.append(' '.join(current_line))
        
        return lines
    
    def create_dynamic_anniversary_card(self, employee_row):
        """Generate enhanced anniversary card with visible split background design"""
        try:
            # Step 1: Create canvas with light gray background
            canvas = Image.new('RGB', self.canvas_size, self.hex_to_rgb(self.colors['background']))
            draw = ImageDraw.Draw(canvas)
            
            width, height = self.canvas_size
            
            # Get fonts
            fonts = self.get_dynamic_fonts()
            
            # Calculate anniversary details
            joining_date = pd.to_datetime(employee_row['JoiningDate'])
            years, ordinal_years, custom_message = self.calculate_anniversary_details(joining_date)
            
            employee_name = str(employee_row['Name']) if pd.notna(employee_row['Name']) else "Unknown"
            employee_dept = str(employee_row['Department']) if pd.notna(employee_row['Department']) else "Unknown Department"
            
            # Define colors
            brand_red = self.hex_to_rgb(self.colors['brand_red'])
            black = self.hex_to_rgb(self.colors['black'])
            white = self.hex_to_rgb(self.colors['white'])
            
            # ===== Building background - SUBTLE and ELEGANT (matching reference template) =====
            print(f"\n   🏢 Loading office building background for upper portion...")
            building_path = r'C:\Projects\AnniversaryAutomation\photos\OfficeBuilding.jpg'

            try:
                if os.path.exists(building_path):
                    building_img = Image.open(building_path)
                    
                    # Convert to RGB
                    if building_img.mode != 'RGB':
                        building_img = building_img.convert('RGB')
                    
                    # Calculate height - cover top area (30% of canvas)
                    upper_height = int(height * 0.30)  # 324px for 1080px canvas
                    
                    # Resize building to canvas width, maintaining aspect ratio
                    aspect_ratio = building_img.height / building_img.width
                    new_width = width
                    new_height = int(new_width * aspect_ratio)
                    
                    building_img = building_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    
                    # Crop to fit the area
                    if new_height > upper_height:
                        building_img = building_img.crop((0, 0, new_width, upper_height))
                    else:
                        upper_height = new_height
                    
                    # === SUBTLE EFFECT - matching reference template ===
                    from PIL import ImageEnhance
                    
                    # 1. Convert to near-grayscale (minimal color saturation)
                    color_enhancer = ImageEnhance.Color(building_img)
                    building_img = color_enhancer.enhance(0.15)  # Almost grayscale, very subtle color
                    
                    # 2. Make it significantly brighter (washed out effect)
                    brightness_enhancer = ImageEnhance.Brightness(building_img)
                    building_img = brightness_enhancer.enhance(1.7)  # Very bright, faded look
                    
                    # 3. Reduce contrast for softer, dimmer appearance
                    contrast_enhancer = ImageEnhance.Contrast(building_img)
                    building_img = contrast_enhancer.enhance(0.6)  # Low contrast, subtle
                    
                    # Create SUBTLE opacity mask with smooth gradient fade
                    mask = Image.new('L', (width, upper_height), 0)
                    mask_draw = ImageDraw.Draw(mask)
                    
                    # Start fading halfway through the building area
                    fade_start = int(upper_height * 0.50)  # Fade starts at 50% height
                    
                    for y in range(upper_height):
                        if y < fade_start:
                            # Top half: Subtle visibility (35% opacity)
                            opacity = int(255 * 0.35)
                        else:
                            # Bottom half: Smooth fade from 35% to 0%
                            fade_progress = (y - fade_start) / (upper_height - fade_start)
                            opacity = int(255 * (0.35 * (1 - fade_progress)))
                        
                        mask_draw.line([(0, y), (width, y)], fill=opacity)
                    
                    # Create full canvas overlay
                    building_layer = Image.new('RGB', self.canvas_size, self.hex_to_rgb(self.colors['background']))
                    building_layer.paste(building_img, (0, 0))
                    
                    # Create mask for full canvas
                    full_mask = Image.new('L', self.canvas_size, 0)
                    full_mask.paste(mask, (0, 0))
                    
                    # Apply the subtle building image
                    canvas = Image.composite(building_layer, canvas, full_mask)
                    
                    print(f"   ✅ Building background applied - SUBTLE & ELEGANT")
                    print(f"   ✅ Coverage: Full width × Top {upper_height}px (30%)")
                    print(f"   ✅ Style: Desaturated, dimmed, smooth fade (matches reference)")
                else:
                    print(f"   ⚠️  Building image not found: {building_path}")
            except Exception as bg_error:
                print(f"   ⚠️  Could not load building background: {bg_error}")
                import traceback
                traceback.print_exc()
                        
            # ===== Add more visible company logo pattern in lower portion =====
            print(f"\n   🎨 Creating more visible company logo background pattern...")
            logo_pattern = self.create_background_pattern(self.canvas_size)
            canvas.paste(logo_pattern, (0, 0), logo_pattern)
            
            # Recreate draw object after pasting
            draw = ImageDraw.Draw(canvas)
            
            # ===== Add decorative corner elements =====
            corner_size = 80
            corner_color = brand_red
            corner_width = 4
            
            # Top-left corner
            draw.line([(30, 30), (30 + corner_size, 30)], fill=corner_color, width=corner_width)
            draw.line([(30, 30), (30, 30 + corner_size)], fill=corner_color, width=corner_width)
            
            # Top-right corner
            draw.line([(width - 30 - corner_size, 30), (width - 30, 30)], fill=corner_color, width=corner_width)
            draw.line([(width - 30, 30), (width - 30, 30 + corner_size)], fill=corner_color, width=corner_width)
            
            # Bottom-left corner
            draw.line([(30, height - 30), (30 + corner_size, height - 30)], fill=corner_color, width=corner_width)
            draw.line([(30, height - 30 - corner_size), (30, height - 30)], fill=corner_color, width=corner_width)
            
            # Bottom-right corner
            draw.line([(width - 30 - corner_size, height - 30), (width - 30, height - 30)], fill=corner_color, width=corner_width)
            draw.line([(width - 30, height - 30 - corner_size), (width - 30, height - 30)], fill=corner_color, width=corner_width)
            
            # Title: "Happy [Xst] Work Anniversary!" in BLACK - centered
            title_y = 120
            anniversary_text = f"Happy {ordinal_years} Work Anniversary!"
            self.draw_centered_text(draw, anniversary_text, title_y, fonts['title'], black, width)
            
            # Add employee photo with enhanced frame
            photo_target_size = (350, 450)
            photo_x = (width - photo_target_size[0]) // 2
            photo_y = 250
            
            print(f"\n   🖼️  Processing photo for {employee_name}...")
            
            # Try to get and process the photo
            employee_photo = None
            if 'PhotoPath' in employee_row and pd.notna(employee_row['PhotoPath']):
                photo_path = str(employee_row['PhotoPath']).strip()
                print(f"   📝 PhotoPath from Excel: '{photo_path}'")
                
                if photo_path:
                    try:
                        employee_photo = self.process_employee_photo(photo_path, photo_target_size)
                        print(f"   ✅ Photo processed successfully!")
                    except Exception as photo_error:
                        print(f"   ⚠️  Could not process photo: {photo_error}")
                        employee_photo = None
                else:
                    print(f"   ⚠️  PhotoPath is empty string")
            else:
                print(f"   ⚠️  No PhotoPath column or value is NaN")
            
            # Use placeholder if photo processing failed
            if employee_photo is None:
                print(f"   🎨 Creating placeholder photo...")
                employee_photo = self.create_placeholder_photo(photo_target_size)
                print(f"   ✅ Placeholder created")
            
            # Create enhanced photo frame with shadow
            shadow_offset = 8
            shadow_color = (180, 180, 180)
            shadow_box = [
                photo_x + shadow_offset,
                photo_y + shadow_offset,
                photo_x + employee_photo.width + shadow_offset,
                photo_y + employee_photo.height + shadow_offset
            ]
            draw.rectangle(shadow_box, fill=shadow_color)
            
            # Create white border frame
            border_thickness = 15
            frame_box = [
                photo_x - border_thickness,
                photo_y - border_thickness,
                photo_x + employee_photo.width + border_thickness,
                photo_y + employee_photo.height + border_thickness
            ]
            draw.rectangle(frame_box, fill=white, outline=brand_red, width=2)
            
            # Paste the photo onto canvas
            try:
                canvas.paste(employee_photo, (photo_x, photo_y))
                print(f"   ✅ Photo pasted onto card")
            except Exception as paste_error:
                print(f"   ❌ Error pasting photo: {paste_error}")
                placeholder = self.create_placeholder_photo(photo_target_size)
                canvas.paste(placeholder, (photo_x, photo_y))
            
            # Recreate draw after pasting photo
            draw = ImageDraw.Draw(canvas)
            
            # Employee Name in RED - centered below photo
            name_y = photo_y + employee_photo.height + 40
            self.draw_centered_text(draw, employee_name, name_y, fonts['name'], brand_red, width)
            
            # Department in BLACK - centered below name
            dept_y = name_y + 75
            self.draw_centered_text(draw, employee_dept, dept_y, fonts['department'], black, width)
            
            # Custom message in BLACK with SCRIPT FONT
            message_y = dept_y + 60
            message_lines = custom_message.split('\n')
            
            # Maximum width for text
            max_text_width = width - 120
            
            all_wrapped_lines = []
            for line in message_lines:
                wrapped = self.wrap_text(line, fonts['message'], max_text_width, draw)
                all_wrapped_lines.extend(wrapped)
            
            # Draw each wrapped line
            line_height = 45
            last_line_y = message_y
            for i, line in enumerate(all_wrapped_lines):
                line_y = message_y + (i * line_height)
                self.draw_centered_text(draw, line, line_y, fonts['message'], black, width)
                last_line_y = line_y
            
            # ===== Move company logo lower to avoid overlap with message =====
            # Calculate position after message text (with extra spacing)
            logo_y = max(last_line_y + 80, height - 150)  # At least 80px below message
            
            company_logo = self.load_company_logo(target_width=200)
            
            if company_logo is not None:
                # Center the logo
                logo_x = (width - company_logo.width) // 2
                
                # Paste logo with transparency support
                if company_logo.mode == 'RGBA':
                    canvas.paste(company_logo, (logo_x, logo_y), company_logo)
                else:
                    canvas.paste(company_logo, (logo_x, logo_y))
                
                print(f"   ✅ Company logo placed at y={logo_y} (no overlap with text)")
            else:
                # Fallback: Draw company initials text if logo not found
                print(f"   ⚠️  Logo not found, using text fallback")
                fallback_font = fonts.get('department', ImageFont.load_default())
                self.draw_centered_text(draw, "CO", logo_y, fallback_font, brand_red, width)
            
            # Save the card
            output_filename = f"anniversary_{employee_name.replace(' ', '_')}_{years}years.png"
            output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_filename)
            
            canvas.save(output_path, quality=95, optimize=True)
            print(f"\n   💾 Card saved: {output_path}")
            
            return output_path
            
        except Exception as e:
            print(f"❌ Error creating anniversary card: {e}")
            self.logger.error(f"Error creating card: {e}")
            raise
            
    
    
    def generate_anniversary_email_content(self, employee_row, years, card_path):
        """Generate email with personalized intro text + anniversary card image"""
        employee_name = str(employee_row['Name'])
        first_name = employee_name.split()[0]  # Extract first name
        ordinal_years = self.get_ordinal_suffix(years)
        
        # Subject line
        subject = f"🎉 Happy {ordinal_years} Work Anniversary, {first_name}!"
        
        # ✅ Adjust wording for 1st year (singular)
        year_phrase = "this year" if years == 1 else "these years"
        
        # Body with 3 intro lines (ALL BOLD) + embedded card image
        body = f"""
        <html>
        <head>
            <style>
                body {{
                    margin: 0;
                    padding: 0;
                    background-color: #FFFFFF;
                    font-family: 'Times New Roman', Times, serif;
                    color: #000000;
                }}
                .email-container {{
                    max-width: 600px;
                    margin: 0 auto;
                    padding: 20px;
                }}
                .greeting {{
                    font-size: 20px;
                    font-weight: bold;  /* ✅ BOLD */
                    margin-bottom: 15px;
                    color: #000000;
                    font-family: 'Times New Roman', Times, serif;
                }}
                .intro-line {{
                    font-size: 20px;
                    font-weight: bold;  /* ✅ BOLD - Changed from normal */
                    line-height: 1.6;
                    margin-bottom: 12px;
                    color: #000000;
                    font-family: 'Times New Roman', Times, serif;
                }}
                .card-container {{
                    text-align: center;
                    margin-top: 25px;
                    padding: 0;
                }}
                .card-image {{
                    width: 450px;
                    height: auto;
                    display: block;
                    margin: 0 auto;
                    border: none;
                }}
                
                /* Responsive for mobile */
                @media only screen and (max-width: 600px) {{
                    .card-image {{
                        width: 100% !important;
                        max-width: 450px;
                    }}
                    .greeting {{
                        font-size: 18px;
                    }}
                    .intro-line {{
                        font-size: 18px;
                    }}
                }}
            </style>
        </head>
        <body>
            <div class="email-container">
                <!-- ✅ 3 INTRODUCTORY LINES - ALL BOLD -->
                <p class="greeting">Hello {first_name}!</p>
                
                <p class="intro-line">Wish you a Happy {ordinal_years} Year Anniversary!</p>
                
                <p class="intro-line">Congratulations on making it through your {self.year_to_text(years)} year, it's been wonderful watching you shine and mark {year_phrase}.</p>
                
                <!-- ✅ ANNIVERSARY CARD IMAGE -->
                <div class="card-container">
                    <img src="cid:anniversary_card" class="card-image" alt="Anniversary Card" width="450">
                </div>
            </div>
        </body>
        </html>
        """
        
        return subject, body


    def year_to_text(self, years):
        """Convert number to ordinal text (e.g., 1 → 'first', 3 → 'third')"""
        text_map = {
            1: 'first',
            2: 'second',
            3: 'third',
            4: 'fourth',
            5: 'fifth',
            6: 'sixth',
            7: 'seventh',
            8: 'eighth',
            9: 'ninth',
            10: 'tenth',
            11: 'eleventh',
            12: 'twelfth',
            13: 'thirteenth',
            14: 'fourteenth',
            15: 'fifteenth',
            20: 'twentieth',
            25: 'twenty-fifth',
            30: 'thirtieth'
        }
        return text_map.get(years, f'{self.get_ordinal_suffix(years).lower()}')


    def send_anniversary_email(self, recipient_email, subject, body, attachment_path, cc_email=None):
        """Send anniversary email with card embedded in body (not as attachment)"""
        try:
            print(f"\n   📧 Initializing Outlook application...")
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            
            print(f"   📝 Setting email properties...")
            mail.To = recipient_email
            mail.Subject = subject
            mail.HTMLBody = body

            # ✅ ADD CC EMAIL IF PROVIDED
            if cc_email and pd.notna(cc_email) and str(cc_email).strip():
                mail.CC = str(cc_email).strip()
                print(f"   📎 CC added: {cc_email}")

            # Embed the card image in the email body
            print(f"   📎 Checking card image: {attachment_path}")
            if os.path.exists(attachment_path):
                abs_path = os.path.abspath(attachment_path)
                print(f"   🖼️  Embedding card image in email body...")
                
                # Add as attachment with Content-ID for inline display
                attachment = mail.Attachments.Add(abs_path)
                
                # Set properties for inline display
                # Property schema: http://schemas.microsoft.com/mapi/proptag/0x3712001F
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", 
                    "anniversary_card"
                )
                
                print(f"   ✅ Card embedded successfully in email body")
            else:
                print(f"   ⚠️  Warning: Card image not found at: {attachment_path}")
                print(f"   📁 Searched in: {os.path.abspath(attachment_path)}")
            
            # Send the email
            print(f"   📤 Sending email to {recipient_email}...")
            mail.Send()
            print(f"   ✅ Email sent successfully to {recipient_email}")
            
            self.logger.info(f"Email sent to {recipient_email}")
            return True
            
        except Exception as e:
            print(f"   ❌ Error sending email to {recipient_email}")
            print(f"   📝 Error details: {type(e).__name__}: {str(e)}")
            self.logger.error(f"Email error for {recipient_email}: {e}")
            
            import traceback
            print(f"   🔍 Full traceback:")
            traceback.print_exc()
            
            return False
    
    def process_anniversaries(self):
        """Main process to generate cards and send emails for all employees"""
        print("\n" + "="*70)
        print("🎉 ANNIVERSARY CARD GENERATOR - STARTED")
        print("="*70)
        
        today = datetime.now()
        current_month = today.month
        current_day = today.day
        
        # Filter employees with anniversaries this month
        self.employee_data['JoiningDate'] = pd.to_datetime(self.employee_data['JoiningDate'])
        anniversary_employees = self.employee_data[
            (self.employee_data['JoiningDate'].dt.month == current_month) &
            (self.employee_data['JoiningDate'].dt.day == current_day)
        ]
        
        if len(anniversary_employees) == 0:
            print(f"\n📅 No anniversaries found for today ({today.strftime('%B %d, %Y')})")
            print("\n" + "="*70)
            return
        
        print(f"\n📅 Found {len(anniversary_employees)} anniversary(ies) for today!")
        print("="*70)
        
        successful = 0
        failed = 0
        
        for idx, employee in anniversary_employees.iterrows():
            try:
                employee_name = str(employee['Name'])
                employee_email = str(employee['Email'])
                
                print(f"\n{'='*70}")
                print(f"👤 Processing: {employee_name}")
                print(f"📧 Email: {employee_email}")
                print(f"{'='*70}")
                
                # Calculate years
                joining_date = pd.to_datetime(employee['JoiningDate'])
                years = today.year - joining_date.year
                if today.month < joining_date.month or (today.month == joining_date.month and today.day < joining_date.day):
                    years -= 1
                
                print(f"📊 Anniversary: {years} year(s)")
                
                # Generate anniversary card
                print(f"\n🎨 Creating anniversary card...")
                card_path = self.create_dynamic_anniversary_card(employee)
                
                if os.path.exists(card_path):
                    print(f"✅ Card created successfully!")
                    
                    # Generate email content (UPDATED - pass card_path)
                    print(f"\n📝 Generating email content...")
                    subject, body = self.generate_anniversary_email_content(employee, years, card_path)
                    
                    # Send email
                    # Send email
                    # Send email
                    print(f"\n📨 Sending email...")

                    # ✅ GET CC EMAIL FROM DATAFRAME - WITH ENHANCED DEBUGGING
                    cc_email = None

                    # Debug: Print all available columns
                    print(f"\n   🔍 DEBUG: Available columns in employee data:")
                    print(f"   📋 Columns: {list(employee.index)}")

                    # Debug: Check if CC column exists
                    if 'CC' in employee.index:
                        print(f"   ✅ CC column FOUND in data")
                        
                        # Debug: Print raw CC value
                        raw_cc = employee['CC']
                        print(f"   📝 Raw CC value: '{raw_cc}' (type: {type(raw_cc).__name__})")
                        
                        # Check if CC has a value
                        if pd.notna(raw_cc):
                            cc_email = str(raw_cc).strip()
                            if cc_email:  # Check if not empty string
                                print(f"   ✅ CC will be added: {cc_email}")
                            else:
                                print(f"   ⚠️  CC column exists but value is empty string")
                        else:
                            print(f"   ⚠️  CC column exists but value is NaN/None")
                    else:
                        print(f"   ❌ CC column NOT FOUND in data")
                        print(f"   💡 Make sure your Excel file has a column named exactly 'CC'")

                    print(f"\n   📧 Final CC value to be used: {cc_email if cc_email else 'None (no CC will be added)'}")

                    if self.send_anniversary_email(employee_email, subject, body, card_path, cc_email):
                        successful += 1
                        print(f"✅ Successfully processed {employee_name}")
                    else:
                        failed += 1
                        print(f"❌ Failed to send email for {employee_name}")
                else:
                    failed += 1
                    print(f"❌ Card file not found for {employee_name}")
                    
            except Exception as e:
                failed += 1
                print(f"❌ Error processing {employee_name}: {e}")
                self.logger.error(f"Error processing {employee_name}: {e}")
        
        # Summary
        print("\n" + "="*70)
        print("📊 PROCESSING SUMMARY")
        print("="*70)
        print(f"✅ Successful: {successful}")
        print(f"❌ Failed: {failed}")
        print(f"📋 Total: {successful + failed}")
        print("="*70)
        print("🎉 ANNIVERSARY CARD GENERATOR - COMPLETED")
        print("="*70 + "\n")


def main():
    """Main execution function"""
    try:
        # Path to your employee data Excel file
        employee_data_file = r'C:\Projects\AnniversaryAutomation\employee_data.xlsx'
        
        print("\n" + "="*70)
        print("🚀 INITIALIZING ANNIVERSARY AUTOMATION")
        print("="*70)
        
        # Check if file exists
        if not os.path.exists(employee_data_file):
            print(f"\n❌ ERROR: Employee data file not found!")
            print(f"📁 Expected location: {employee_data_file}")
            print("\nPlease ensure the file exists at the specified location.")
            return
        
        print(f"\n✅ Employee data file found: {employee_data_file}")
        
        # Initialize the generator
        print(f"\n🔧 Initializing Anniversary Generator...")
        generator = DynamicAnniversaryGenerator(employee_data_file)
        
        print(f"✅ Generator initialized successfully!")
        
        # Process anniversaries
        generator.process_anniversaries()
        
    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {e}")
        logging.error(f"Critical error in main: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()