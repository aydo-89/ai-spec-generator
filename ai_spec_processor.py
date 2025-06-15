import re
import json
import difflib
import pathlib
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, List, Optional, Tuple
import openai
from dataclasses import dataclass
import logging
from io import BytesIO

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@dataclass
class MaterialMatch:
    original: str
    standardized: str
    confidence: float
    method: str  # 'exact', 'fuzzy', 'ai'
    reasoning: Optional[str] = None

@dataclass
class ProcessingResult:
    success: bool
    samples_processed: int
    total_samples: int
    matches_by_method: Dict[str, int]
    errors: List[str]
    output_file: Optional[bytes] = None

class AISpecProcessor:
    def __init__(self, api_key: str):
        """Initialize the AI-enhanced spec processor"""
        self.client = openai.OpenAI(api_key=api_key)
        self.ai_cache = {}  # Cache AI results to avoid repeated API calls
        self.materials = []
        self.exact_lookup = {}
        
        # Template field mappings based on analysis
        self.template_mapping = {
            'season_gender': ('B', 1),
            'sample_name': ('G', 3),
            'color_name': ('L', 3),
            'factory_ref': ('B', 5),
            'sample_number': ('B', 26),
            'upper': ('H', 6),
            'trim': ('H', 8),
            'lining': ('H', 9),
            'sock_topcover': ('H', 10),
            'sock_label': ('H', 11),
            'insole': ('H', 12),
            'midsole': ('H', 13),
            'outsole': ('H', 14),
            'outsole_treatment': ('H', 15),
            'detail_stitching': ('H', 16),
            'reg_stitching': ('H', 17),
            'hardware': ('H', 18),
            'other': ('H', 19)
        }
        
    def load_bom(self, bom_file: BytesIO) -> bool:
        """Load and process the BOM file"""
        try:
            logger.info("Loading BOM file...")
            raw_bom = pd.read_excel(bom_file)
            
            # Extract materials from BOM structure (category: material format)
            materials = []
            for idx, row in raw_bom.iterrows():
                for col_val in row:
                    if pd.notna(col_val) and isinstance(col_val, str):
                        # Look for "Category: Material" format
                        if ':' in col_val and len(col_val.split(':', 1)) == 2:
                            category, material = col_val.split(':', 1)
                            material = material.strip()
                            if material and len(material) > 2:
                                materials.append(material)
                        elif len(col_val.strip()) > 2:
                            materials.append(col_val.strip())
            
            # Clean and deduplicate materials
            self.materials = list(set([m for m in materials if len(m) > 2]))
            self.exact_lookup = {m.lower(): m for m in self.materials}
            
            logger.info(f"Loaded {len(self.materials)} materials from BOM")
            return True
            
        except Exception as e:
            logger.error(f"Error loading BOM: {e}")
            return False
    
    def parse_vendor_materials(self, upper_text: str, sole_text: str) -> Dict[str, str]:
        """AI-enhanced parsing of vendor material descriptions with intelligent categorization"""
        if not isinstance(upper_text, str):
            upper_text = ""
        if not isinstance(sole_text, str):
            sole_text = ""
            
        try:
            # Combine all material text for AI analysis
            combined_text = f"Upper materials: {upper_text}\nSole materials: {sole_text}"
            
            prompt = f"""You are a footwear material categorization expert. Parse these vendor material descriptions and categorize them into the correct spec sheet fields.

VENDOR MATERIALS:
{combined_text}

SPEC SHEET CATEGORIES:
- Upper: Main upper materials
- Trim: Decorative elements, laces, bindings
- Lining: Interior lining materials  
- Sock (topcover): Footbed/sock liner materials
- Sock Label: Sock labels or sock printing
- Insole: Insole materials
- Midsole: Midsole materials
- Outsole: Outsole/bottom materials
- Outsole Treatment: Special outsole treatments
- Detail Stitching: Decorative or contrast stitching
- Reg Stitching: Regular/standard stitching
- Hardware: Buckles, eyelets, metal components
- Other: Anything that doesn't fit above categories

VENDOR TERMINOLOGY VARIATIONS:
- "Ball stitching", "Small stitching" → Detail Stitching or Reg Stitching
- "Pile", "Faux fur" → Upper or Lining
- "TPR", "EVA", "Rubber" → Outsole
- "Microfiber", "Berber" → Lining
- Material codes like "VP-052", "W211" → Extract the material name

RULES:
1. Parse colon-delimited entries (e.g., "Upper: Brown Suede")
2. Handle vendor-specific terminology intelligently
3. If unsure about category, put in "Other" with format "Label: Material"
4. Extract clean material names (remove codes when possible)
5. Never leave materials uncategorized

Return ONLY valid JSON:
{{
    "Upper": "material name or empty string",
    "Trim": "material name or empty string", 
    "Lining": "material name or empty string",
    "Sock (topcover)": "material name or empty string",
    "Sock Label": "material name or empty string",
    "Insole": "material name or empty string",
    "Midsole": "material name or empty string", 
    "Outsole": "material name or empty string",
    "Outsole Treatment": "material name or empty string",
    "Detail Stitching": "material name or empty string",
    "Reg Stitching": "material name or empty string",
    "Hardware": "material name or empty string",
    "Other": "Label: Material (if any uncategorized items)"
}}"""

            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=500
            )
            
            result_text = response.choices[0].message.content.strip()
            if result_text.startswith('```json'):
                result_text = result_text.replace('```json', '').replace('```', '').strip()
            
            parsed_materials = json.loads(result_text)
            
            # Clean up empty values
            return {k: v for k, v in parsed_materials.items() if v and v.strip()}
            
        except Exception as e:
            logger.warning(f"AI parsing error: {e}")
            # Fallback to regex parsing
            return self._regex_parse_materials(upper_text, sole_text)
    
    def _regex_parse_materials(self, upper_text: str, sole_text: str) -> Dict[str, str]:
        """Fallback regex parsing for vendor materials"""
        materials = {}
        
        # Simple patterns for common formats
        patterns = [
            (r'upper[^:]*?:\s*([^,\n]+)', 'Upper'),
            (r'lining[^:]*?:\s*([^,\n]+)', 'Lining'),
            (r'sole[^:]*?:\s*([^,\n]+)', 'Outsole'),
            (r'outsole[^:]*?:\s*([^,\n]+)', 'Outsole'),
            (r'midsole[^:]*?:\s*([^,\n]+)', 'Midsole'),
        ]
        
        combined_text = f"{upper_text} {sole_text}"
        
        for pattern, category in patterns:
            matches = re.finditer(pattern, combined_text, re.IGNORECASE)
            for match in matches:
                material = match.group(1).strip()
                if material and len(material) > 2:
                    materials[category] = material
                    break
        
        # If no patterns match, put raw materials in appropriate categories
        if not materials:
            if upper_text.strip():
                materials['Upper'] = upper_text.strip()
            if sole_text.strip():
                materials['Outsole'] = sole_text.strip()
        
        return materials
    
    def standardize_material(self, raw_material: str) -> MaterialMatch:
        """Enhanced material standardization with BOM cross-reference"""
        if not isinstance(raw_material, str) or not raw_material.strip():
            return MaterialMatch(raw_material, raw_material, 0.0, 'no')
        
        clean_material = raw_material.strip()
        material_lower = clean_material.lower()
        
        # 1. Try exact match first
        if material_lower in self.exact_lookup:
            return MaterialMatch(
                original=raw_material,
                standardized=self.exact_lookup[material_lower],
                confidence=1.0,
                method='exact'
            )
        
        # 2. Try fuzzy matching
        fuzzy_matches = difflib.get_close_matches(
            material_lower, list(self.exact_lookup.keys()), n=1, cutoff=0.75
        )
        if fuzzy_matches:
            return MaterialMatch(
                original=raw_material,
                standardized=self.exact_lookup[fuzzy_matches[0]],
                confidence=0.85,
                method='fuzzy'
            )
        
        # 3. AI-enhanced matching for partial names
        ai_result = self._ai_material_match(clean_material)
        if ai_result:
            return ai_result
        
        # 4. No match found - return original
        return MaterialMatch(raw_material, raw_material, 0.0, 'no')
    
    def _ai_material_match(self, material: str) -> Optional[MaterialMatch]:
        """Use AI to find the best BOM material match for partial descriptions"""
        if material in self.ai_cache:
            return self.ai_cache[material]
        
        try:
            # Create prompt with BOM materials
            bom_list = "\n".join([f"- {mat}" for mat in self.materials[:30]])
            
            prompt = f"""You are a footwear material matching expert. Find the best match for this vendor material description.

VENDOR MATERIAL: "{material}"

STANDARDIZED BOM MATERIALS:
{bom_list}

MATCHING RULES:
- Look for partial matches (e.g., "Brown Suede" → "W1063 Minnetonka Brown Cow Suede")
- Consider synonyms (e.g., "Pile" → "Faux Fur", "TPR" → "Rubber")
- Match colors and material types
- Only return exact materials from the BOM list above
- Confidence must be 0.7+ for a valid match

Return ONLY valid JSON:
{{
    "best_match": "exact material name from BOM or null if no good match",
    "confidence": 0.0-1.0,
    "reasoning": "brief explanation of the match"
}}"""

            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=200
            )
            
            result_text = response.choices[0].message.content.strip()
            if result_text.startswith('```json'):
                result_text = result_text.replace('```json', '').replace('```', '').strip()
            
            result = json.loads(result_text)
            
            if result.get('best_match') and result.get('confidence', 0) >= 0.7:
                match = MaterialMatch(
                    original=material,
                    standardized=result['best_match'],
                    confidence=result['confidence'],
                    method='ai',
                    reasoning=result.get('reasoning')
                )
                self.ai_cache[material] = match
                return match
            
        except Exception as e:
            logger.warning(f"AI matching error for '{material}': {e}")
        
        return None
    
    def infer_color_name(self, materials_dict: Dict[str, str]) -> str:
        """AI-powered color inference from material descriptions"""
        try:
            # Combine all materials for color analysis
            all_materials = " ".join([f"{k}: {v}" for k, v in materials_dict.items() if v])
            
            prompt = f"""Extract the primary color name from these footwear materials:

MATERIALS: {all_materials}

RULES:
- Return the main/primary color (e.g., "Brown", "Black", "Natural", "Cognac")
- Look in Upper materials first, then other materials
- For multi-color items, pick the dominant color
- Use standard color names (not codes)
- If no clear color, return "Natural"

Return ONLY the color name (one or two words max):"""

            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=50
            )
            
            color = response.choices[0].message.content.strip().strip('"')
            return color if color else "Natural"
            
        except Exception as e:
            logger.warning(f"Color inference error: {e}")
            return "Natural"
    
    def process_spec_sheets(self, dev_log_file: BytesIO, template_file: BytesIO) -> ProcessingResult:
        """Main processing function with enhanced material categorization"""
        try:
            logger.info("Starting spec sheet processing...")
            
            # Load development log
            dev = pd.read_excel(dev_log_file)
            dev.columns = dev.iloc[0]  # Promote first row to header
            dev = dev.drop(index=0).reset_index(drop=True)
            
            logger.info(f"Loaded {len(dev)} samples from development log")
            
            # Load template
            wb_master = load_workbook(template_file)
            base_sheet = wb_master.active
            
            # Process each sample
            stats = {
                'exact_matches': 0,
                'fuzzy_matches': 0,
                'ai_matches': 0,
                'no_matches': 0
            }
            
            errors = []
            processed_count = 0
            
            for idx, row in dev.iterrows():
                try:
                    sample_name = str(row.get('Sample Name', f'Sample_{idx}'))
                    logger.info(f"Processing: {sample_name}")
                    
                    # Create new sheet
                    sheet = wb_master.copy_worksheet(base_sheet)
                    clean_name = re.sub(r'[\\/*?[\]:]+', '_', sample_name)
                    sheet.title = (clean_name[:28] + '...') if len(clean_name) > 31 else clean_name
                    
                    # Fill basic metadata
                    season = str(row.get('Season', ''))
                    gender = str(row.get('Gender', ''))
                    season_gender = f"{season}, {gender}".strip(', ')
                    
                    sheet['B1'] = season_gender  # Season & Gender
                    sheet['G3'] = sample_name    # Sample Name
                    sheet['B5'] = str(row.get('Factory ref #', ''))  # Factory Reference
                    sheet['B26'] = str(row.get('Sample Order No.', ''))  # Sample Number
                    
                    # Parse vendor materials with AI categorization
                    upper_text = str(row.get('Upper', ''))
                    sole_text = str(row.get('Sole (ref # only)', '')) or str(row.get('Sole', ''))
                    
                    parsed_materials = self.parse_vendor_materials(upper_text, sole_text)
                    
                    # Infer color name from materials
                    color_name = self.infer_color_name(parsed_materials)
                    sheet['L3'] = color_name  # Color Name
                    
                    # Process and standardize each material
                    for category, raw_material in parsed_materials.items():
                        if raw_material and raw_material.strip():
                            # Standardize against BOM
                            match_result = self.standardize_material(raw_material)
                            
                            # Update stats
                            method_key = f'{match_result.method}_matches'
                            if method_key in stats:
                                stats[method_key] += 1
                            else:
                                stats['no_matches'] += 1
                            
                            # Map category to template field
                            category_key = category.lower().replace(' ', '_').replace('(', '').replace(')', '')
                            if category_key in self.template_mapping:
                                col, row_num = self.template_mapping[category_key]
                                sheet[f'{col}{row_num}'] = match_result.standardized
                            else:
                                # Put in "Other" category with label
                                sheet['H19'] = f"{category}: {match_result.standardized}"
                    
                    processed_count += 1
                    
                except Exception as e:
                    error_msg = f"Error processing {sample_name}: {str(e)}"
                    errors.append(error_msg)
                    logger.error(error_msg)
            
            # Remove original template sheet
            wb_master.remove(base_sheet)
            
            # Save to BytesIO
            output_buffer = BytesIO()
            wb_master.save(output_buffer)
            output_buffer.seek(0)
            
            return ProcessingResult(
                success=True,
                samples_processed=processed_count,
                total_samples=len(dev),
                matches_by_method=stats,
                errors=errors,
                output_file=output_buffer.getvalue()
            )
            
        except Exception as e:
            logger.error(f"Processing failed: {e}")
            return ProcessingResult(
                success=False,
                samples_processed=0,
                total_samples=0,
                matches_by_method={},
                errors=[str(e)]
            ) 