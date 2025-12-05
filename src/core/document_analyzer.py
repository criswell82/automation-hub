"""
Document Structure Analyzer

Analyzes Word documents to determine optimal template conversion approach.
Detects placeholders, infers parameters, and recommends template modes.
"""
import re
import logging
from enum import Enum
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Tuple
from pathlib import Path

from modules.word_automation.document_handler import DocumentHandler


class TemplateMode(Enum):
    """Template generation modes."""
    FILL_IN = "fill_in"      # Document has placeholders to replace
    GENERATE = "generate"    # Document structure to recreate from scratch
    CONTENT = "content"      # Static document to reuse as-is
    PATTERN = "pattern"      # Repeating structure (like mail merge)


@dataclass
class PlaceholderInfo:
    """Information about a detected placeholder."""
    name: str                      # Variable name (e.g., 'client_name')
    pattern: str                   # Full pattern (e.g., '{{client_name}}')
    suggested_type: str            # Inferred parameter type
    locations: List[str] = field(default_factory=list)  # Where found
    suggested_description: str = ""  # Auto-generated description
    count: int = 1                 # Number of occurrences


class DocumentAnalyzer:
    """
    Analyze documents to determine optimal template conversion strategy.

    Capabilities:
    - Extract document structure (headings, tables, content)
    - Detect placeholders in multiple formats
    - Infer parameter types from placeholder names
    - Recommend template mode (fill-in, generate, content, pattern)
    - Generate WORKFLOW_META parameters automatically
    """

    # Placeholder pattern regexes (in order of preference)
    PLACEHOLDER_PATTERNS = [
        (r'\{\{(\w+)\}\}', 'jinja2'),        # {{variable}}
        (r'\[([A-Z_][A-Z0-9_]*)\]', 'bracket'),  # [FIELD]
        (r'\{(\w+)\}', 'brace'),             # {placeholder}
        (r'<(\w+)>', 'angle'),               # <TOKEN>
        (r'__(\w+)__', 'underscore'),        # __VAR__
    ]

    # Type inference patterns
    TYPE_PATTERNS = {
        'date': [r'.*_date$', r'^date_.*', r'.*_on$', r'.*_at$'],
        'email': [r'.*_email$', r'^email_.*', r'.*_address$'],
        'number': [r'.*_(amount|price|cost|qty|quantity|count|total)$', r'^(amount|price|cost|qty|quantity|count|total)_.*'],
        'file': [r'.*_(file|path|document|doc)$', r'^(file|path|document|doc)_.*'],
        'boolean': [r'^is_.*', r'^has_.*', r'.*_enabled$', r'.*_active$'],
        'text': [r'.*_(description|comment|note|remarks?)$'],
    }

    def __init__(self):
        """Initialize the document analyzer."""
        self.logger = logging.getLogger(__name__)

    def analyze_word_document(self, filepath: str) -> Dict[str, Any]:
        """
        Comprehensively analyze a Word document.

        Args:
            filepath: Path to .docx file

        Returns:
            Dictionary containing:
            - mode: Recommended template mode
            - confidence: Confidence score (0-1)
            - structure: Document structure details
            - placeholders: List of detected placeholders
            - parameters: Generated WORKFLOW_META parameters
            - recommended_template_name: Suggested template name
            - recommended_category: Suggested category
        """
        self.logger.info(f"Analyzing document: {filepath}")

        # Load document and extract basic structure
        handler = DocumentHandler(filepath)
        structure = handler.extract_structure()

        # Enhance structure with additional analysis
        structure['complexity_score'] = self._calculate_complexity(structure)
        structure['has_images'] = False  # TODO: Implement image detection
        structure['formatting_variety'] = 'medium'  # TODO: Implement style analysis

        # Detect placeholders
        placeholders = self._detect_all_placeholders(structure)

        # Detect template mode
        mode, confidence = self._detect_template_mode(structure, placeholders)

        # Extract parameters for WORKFLOW_META
        parameters = self._extract_parameters(placeholders, mode)

        # Generate recommendations
        template_name = self._suggest_template_name(filepath, structure)
        category = self._suggest_category(structure, mode)

        analysis = {
            'mode': mode.value,
            'confidence': confidence,
            'structure': structure,
            'placeholders': [self._placeholder_to_dict(p) for p in placeholders],
            'parameters': parameters,
            'recommended_template_name': template_name,
            'recommended_category': category,
            'source_file': str(filepath)
        }

        self.logger.info(f"Analysis complete: mode={mode.value}, "
                        f"{len(placeholders)} placeholders, "
                        f"confidence={confidence:.2f}")

        return analysis

    def _detect_all_placeholders(self, structure: Dict) -> List[PlaceholderInfo]:
        """
        Detect placeholders using multiple pattern styles.

        Args:
            structure: Document structure from extract_structure()

        Returns:
            List of PlaceholderInfo objects
        """
        placeholders_dict = {}  # name -> PlaceholderInfo

        # Search in paragraphs
        for idx, para_text in enumerate(structure.get('paragraphs', [])):
            for pattern, pattern_name in self.PLACEHOLDER_PATTERNS:
                matches = re.finditer(pattern, para_text)
                for match in matches:
                    var_name = match.group(1).lower()
                    full_pattern = match.group(0)

                    if var_name in placeholders_dict:
                        placeholders_dict[var_name].count += 1
                        placeholders_dict[var_name].locations.append(f'paragraph_{idx}')
                    else:
                        placeholders_dict[var_name] = PlaceholderInfo(
                            name=var_name,
                            pattern=full_pattern,
                            suggested_type=self._infer_type(var_name),
                            locations=[f'paragraph_{idx}'],
                            suggested_description=self._generate_description(var_name),
                            count=1
                        )

        # Search in tables
        for table_idx, table_data in enumerate(structure.get('tables', [])):
            for row_idx, row in enumerate(table_data.get('data', [])):
                for col_idx, cell_text in enumerate(row):
                    for pattern, pattern_name in self.PLACEHOLDER_PATTERNS:
                        matches = re.finditer(pattern, cell_text)
                        for match in matches:
                            var_name = match.group(1).lower()
                            full_pattern = match.group(0)
                            location = f'table_{table_idx}_r{row_idx}_c{col_idx}'

                            if var_name in placeholders_dict:
                                placeholders_dict[var_name].count += 1
                                placeholders_dict[var_name].locations.append(location)
                            else:
                                placeholders_dict[var_name] = PlaceholderInfo(
                                    name=var_name,
                                    pattern=full_pattern,
                                    suggested_type=self._infer_type(var_name),
                                    locations=[location],
                                    suggested_description=self._generate_description(var_name),
                                    count=1
                                )

        placeholders = list(placeholders_dict.values())
        self.logger.debug(f"Detected {len(placeholders)} unique placeholders")

        return placeholders

    def _infer_type(self, var_name: str) -> str:
        """
        Infer parameter type from variable name.

        Args:
            var_name: Variable name (e.g., 'client_name', 'invoice_date')

        Returns:
            Type string ('string', 'date', 'number', 'boolean', etc.)
        """
        var_name_lower = var_name.lower()

        # Check each type pattern
        for type_name, patterns in self.TYPE_PATTERNS.items():
            for pattern in patterns:
                if re.match(pattern, var_name_lower):
                    return type_name

        # Default to string
        return 'string'

    def _generate_description(self, var_name: str) -> str:
        """
        Generate human-readable description from variable name.

        Args:
            var_name: Variable name in snake_case

        Returns:
            Description string
        """
        # Convert snake_case to Title Case
        words = var_name.replace('_', ' ').split()
        readable = ' '.join(word.capitalize() for word in words)

        # Add context based on type
        var_type = self._infer_type(var_name)

        if var_type == 'date':
            return f"{readable} (date format)"
        elif var_type == 'email':
            return f"{readable} email address"
        elif var_type == 'number':
            return f"{readable} value"
        elif var_type == 'boolean':
            return f"Whether {readable.lower()}"
        elif var_type == 'file':
            return f"Path to {readable.lower()}"
        else:
            return readable

    def _calculate_complexity(self, structure: Dict) -> float:
        """
        Calculate document complexity score.

        Args:
            structure: Document structure

        Returns:
            Complexity score (0-10)
        """
        score = 0.0

        # Paragraph count (0-2 points)
        para_count = structure.get('total_paragraphs', 0)
        score += min(2.0, para_count / 10)

        # Heading count (0-2 points)
        heading_count = len(structure.get('headings', []))
        score += min(2.0, heading_count / 3)

        # Table count (0-3 points)
        table_count = structure.get('total_tables', 0)
        score += min(3.0, table_count * 1.5)

        # Table complexity (0-2 points)
        for table in structure.get('tables', []):
            cells = table.get('rows', 0) * table.get('cols', 0)
            score += min(2.0, cells / 20)

        # Placeholder count (0-1 point)
        placeholder_count = len(structure.get('placeholders', []))
        score += min(1.0, placeholder_count / 5)

        return round(min(10.0, score), 2)

    def _detect_template_mode(self, structure: Dict, placeholders: List[PlaceholderInfo]) -> Tuple[TemplateMode, float]:
        """
        Detect the optimal template mode.

        Args:
            structure: Document structure
            placeholders: Detected placeholders

        Returns:
            Tuple of (TemplateMode, confidence_score)
        """
        complexity = structure.get('complexity_score', 0)
        placeholder_count = len(placeholders)
        paragraph_count = structure.get('total_paragraphs', 0)
        table_count = structure.get('total_tables', 0)

        # FILL_IN mode: Has placeholders
        if placeholder_count > 0:
            confidence = min(0.95, 0.7 + (placeholder_count * 0.05))
            return TemplateMode.FILL_IN, confidence

        # PATTERN mode: Repeating structure (multiple similar tables)
        if table_count > 1:
            # Check if tables have similar structure
            tables = structure.get('tables', [])
            if len(tables) >= 2:
                # Simple heuristic: similar column counts
                col_counts = [t.get('cols', 0) for t in tables]
                if len(set(col_counts)) <= 2:  # At most 2 different column counts
                    return TemplateMode.PATTERN, 0.75

        # GENERATE mode: Complex structure worth recreating programmatically
        if complexity >= 4.0:
            confidence = min(0.90, 0.5 + (complexity / 20))
            return TemplateMode.GENERATE, confidence

        # CONTENT mode: Simple, static document
        confidence = 0.80
        return TemplateMode.CONTENT, confidence

    def _extract_parameters(self, placeholders: List[PlaceholderInfo], mode: TemplateMode) -> Dict[str, Dict]:
        """
        Generate WORKFLOW_META parameters from placeholders.

        Args:
            placeholders: Detected placeholders
            mode: Template mode

        Returns:
            Parameters dictionary for WORKFLOW_META
        """
        if mode != TemplateMode.FILL_IN or not placeholders:
            # Non-fill-in modes don't need placeholder parameters
            return self._get_default_parameters(mode)

        parameters = {}

        for placeholder in placeholders:
            param_def = {
                'type': placeholder.suggested_type,
                'description': placeholder.suggested_description,
                'required': True,  # Assume all placeholders are required
            }

            # Add type-specific settings
            if placeholder.suggested_type == 'boolean':
                param_def['default'] = False
            elif placeholder.suggested_type == 'number':
                param_def['default'] = 0
            else:
                param_def['default'] = None

            # Add to parameters
            parameters[placeholder.name] = param_def

        # Add common output parameter
        parameters['output_file'] = {
            'type': 'file',
            'description': 'Output path for generated document',
            'required': True,
            'default': None
        }

        return parameters

    def _get_default_parameters(self, mode: TemplateMode) -> Dict[str, Dict]:
        """Get default parameters for non-fill-in modes."""
        if mode == TemplateMode.GENERATE:
            return {
                'output_file': {
                    'type': 'file',
                    'description': 'Output path for generated document',
                    'required': True
                },
                'data': {
                    'type': 'text',
                    'description': 'Data to include in document',
                    'required': False,
                    'default': ''
                }
            }
        elif mode == TemplateMode.PATTERN:
            return {
                'data_source': {
                    'type': 'file',
                    'description': 'CSV or Excel file with data',
                    'required': True
                },
                'output_folder': {
                    'type': 'file',
                    'description': 'Folder for generated documents',
                    'required': True
                }
            }
        elif mode == TemplateMode.CONTENT:
            return {
                'output_file': {
                    'type': 'file',
                    'description': 'Where to copy the document',
                    'required': True
                }
            }

        return {}

    def _suggest_template_name(self, filepath: str, structure: Dict) -> str:
        """
        Suggest a template name based on file name and structure.

        Args:
            filepath: Original file path
            structure: Document structure

        Returns:
            Suggested template name
        """
        # Start with filename
        filename = Path(filepath).stem

        # Clean up common template indicators
        filename = re.sub(r'(template|_template|\.template)', '', filename, flags=re.IGNORECASE)
        filename = re.sub(r'(sample|_sample|\.sample)', '', filename, flags=re.IGNORECASE)

        # Convert to Title Case
        name_parts = re.split(r'[_\-\s]+', filename)
        name = ' '.join(word.capitalize() for word in name_parts if word)

        # If document has first heading, use it if more descriptive
        headings = structure.get('headings', [])
        if headings and len(headings[0]['text']) > len(name):
            name = headings[0]['text']

        # Ensure it doesn't end with common suffixes
        name = re.sub(r'\s+(template|document|doc|file)$', '', name, flags=re.IGNORECASE)

        # Add "Generator" if it's a fill-in template
        if structure.get('placeholders'):
            if not name.endswith('Generator'):
                name += ' Generator'

        return name or 'Document Template'

    def _suggest_category(self, structure: Dict, mode: TemplateMode) -> str:
        """
        Suggest template category based on structure and mode.

        Args:
            structure: Document structure
            mode: Template mode

        Returns:
            Category string
        """
        # Check headings for keywords
        headings_text = ' '.join(h['text'].lower() for h in structure.get('headings', []))

        if any(word in headings_text for word in ['invoice', 'receipt', 'bill', 'statement']):
            return 'Reports'
        elif any(word in headings_text for word in ['report', 'summary', 'analysis']):
            return 'Reports'
        elif any(word in headings_text for word in ['letter', 'memo', 'notice']):
            return 'Email'
        elif any(word in headings_text for word in ['certificate', 'award', 'diploma']):
            return 'Files'

        # Default based on mode
        if mode == TemplateMode.FILL_IN:
            return 'Reports'
        elif mode == TemplateMode.PATTERN:
            return 'Files'
        else:
            return 'Custom'

    def _placeholder_to_dict(self, placeholder: PlaceholderInfo) -> Dict:
        """Convert PlaceholderInfo to dictionary."""
        return {
            'name': placeholder.name,
            'pattern': placeholder.pattern,
            'type': placeholder.suggested_type,
            'locations': placeholder.locations,
            'description': placeholder.suggested_description,
            'count': placeholder.count
        }

    def analyze_email(self, email_source: Any) -> Dict[str, Any]:
        """
        Analyze an email message structure.

        Args:
            email_source: Email source (file path or message object)

        Returns:
            Analysis dictionary

        Note:
            This is a placeholder for future email analysis implementation.
        """
        self.logger.warning("Email analysis not yet implemented")
        return {
            'mode': 'generate',
            'confidence': 0.5,
            'structure': {},
            'placeholders': [],
            'parameters': {},
            'recommended_template_name': 'Email Template',
            'recommended_category': 'Email'
        }
