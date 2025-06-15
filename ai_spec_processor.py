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
        
    def load_bom(self, bom_file: BytesIO) -> bool:
        """Load and process the BOM file"""
        try:
            logger.info("Loading BOM file...")
            raw_bom = pd.read_excel(bom_file, header=None)
            
            # Extract materials from the BOM structure
            materials = []
            for idx, row in raw_bom.iterrows():
                for col_val in row:
                    if pd.notna(col_val) and isinstance(col_val, str):
                        # Look for material names (skip part names with colons)
                        if ':' in col_val:
                            # Split "Part: Material" format
                            parts = col_val.split(':', 1)
                            if len(parts) == 2 and parts[1].strip():
                                materials.append(parts[1].strip())
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
    
    def standardize_material(self, raw_material: str) -> MaterialMatch:
        """Enhanced material standardization with AI"""
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
            material_lower, list(self.exact_lookup.keys()), n=1, cutoff=0.85
        )
        if fuzzy_matches:
            return MaterialMatch(
                original=raw_material,
                standardized=self.exact_lookup[fuzzy_matches[0]],
                confidence=0.85,
                method='fuzzy'
            )
        
        # 3. AI-enhanced matching
        ai_result = self._ai_material_match(clean_material)
        if ai_result:
            return ai_result
        
        # 4. No match found
        return MaterialMatch(raw_material, raw_material, 0.0, 'no')
    
    def _ai_material_match(self, material: str) -> Optional[MaterialMatch]:
        """Use AI to find the best material match"""
        if material in self.ai_cache:
            return self.ai_cache[material]
        
        try:
            # Create prompt with BOM materials
            bom_list = "\n".join([f"- {mat}" for mat in self.materials[:50]])
            
            prompt = f"""You are a footwear material standardization expert.

Given this supplier material description: "{material}"

Find the best match from our standardized BOM materials:
{bom_list}

Return ONLY valid JSON:
{{
    "best_match": "exact material name from BOM or null if no good match",
    "confidence": 0.0-1.0,
    "reasoning": "brief explanation of why this is the best match"
}}

Requirements:
- Only return materials that exist exactly in the BOM list above
- Confidence should be 0.7+ for a match
- Consider synonyms, abbreviations, and common footwear material variations
- If no good match (confidence < 0.7), return null for best_match"""

            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",  # Using mini for cost efficiency
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=200
            )
            
            result_text = response.choices[0].message.content.strip()
            # Remove markdown formatting if present
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
    
    def parse_material_block(self, text: str) -> Dict[str, str]:
        """AI-enhanced parsing of complex material descriptions"""
        if not isinstance(text, str) or not text.strip():
            return {}
        
        try:
            prompt = f"""Extract shoe part materials from this text: "{text}"

Common shoe parts: Upper, Lining, Insole, Outsole, Midsole, Footbed, Heel, Toe, Quarter, Vamp, Sock, Topcover

Return ONLY valid JSON with part names as keys and materials as values:
{{
    "Upper": "material name",
    "Lining": "material name",
    "Outsole": "material name"
}}

Rules:
- Only include parts that are clearly mentioned
- Clean up material names (remove extra spaces, colors, codes)
- Use standard part names
- Extract the core material name (e.g., "Cow Suede" not "W1063 Minnetonka Brown Cow Suede")
- If unclear or empty, return {{}}"""

            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=300
            )
            
            result_text = response.choices[0].message.content.strip()
            if result_text.startswith('```json'):
                result_text = result_text.replace('```json', '').replace('```', '').strip()
            
            return json.loads(result_text)
            
        except Exception as e:
            logger.warning(f"AI parsing error: {e}")
            # Fallback to regex parsing
            return self._regex_parse_materials(text)
    
    def _regex_parse_materials(self, text: str) -> Dict[str, str]:
        """Fallback regex parsing"""
        materials = {}
        
        patterns = [
            r'(Upper[^:]*?):\s*([^-\n]+?)(?:\s*[-\n]|$)',
            r'(Lining[^:]*?):\s*([^,\n]+?)(?:\s*[,\n]|$)',
            r'(Sole[^:]*?):\s*([^,\n]+?)(?:\s*[,\n]|$)',
            r'(Midsole[^:]*?):\s*([^,\n]+?)(?:\s*[,\n]|$)',
            r'(Outsole[^:]*?):\s*([^,\n]+?)(?:\s*[,\n]|$)',
            r'(Footbed[^:]*?):\s*([^,\n]+?)(?:\s*[,\n]|$)',
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                part = match.group(1).strip()
                material = match.group(2).strip()
                
                # Clean up
                part = re.sub(r'[/\s]+', ' ', part.strip())
                material = re.split(r'[,\-]\s*(?:Color|Colour)', material)[0].strip()
                
                if part and material and len(material) > 2:
                    materials[part] = material
        
        return materials
    
    def process_spec_sheets(self, dev_log_file: BytesIO, template_file: BytesIO) -> ProcessingResult:
        """Main processing function"""
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
                    
                    # Fill metadata
                    season_gender = f"{row.get('Season', '')}, {row.get('Gender', '')}"
                    sheet['B1'] = season_gender
                    sheet['C2'] = sample_name
                    sheet['A4'] = row.get('Factory ref #', '') or row.get('Factory Ref #', '')
                    sheet['E2'] = row.get('Sample Order No.', '')
                    
                    # Parse and standardize materials
                    upper_materials = self.parse_material_block(str(row.get('Upper', '')))
                    sole_text = str(row.get('Sole (ref # only)', '')) or str(row.get('Sole', ''))
                    sole_materials = self.parse_material_block(sole_text)
                    
                    all_materials = {**upper_materials, **sole_materials}
                    
                    # Match and standardize materials
                    for part, raw_material in all_materials.items():
                        match_result = self.standardize_material(raw_material)
                        # Update stats safely
                        method_key = f'{match_result.method}_matches'
                        if method_key in stats:
                            stats[method_key] += 1
                        else:
                            stats['no_matches'] += 1
                        
                        # Find template row and fill
                        template_row = self._find_template_row(sheet, part)
                        if template_row:
                            sheet[f'B{template_row}'] = match_result.standardized
                    
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
    
    def _find_template_row(self, sheet: Worksheet, part_name: str) -> Optional[int]:
        """Find where to place a material in the template"""
        for r in range(1, sheet.max_row + 1):
            cell_value = sheet[f'A{r}'].value
            if cell_value and part_name.lower() in str(cell_value).lower():
                return r
        return None 