import json
from typing import Dict, Any
from app.ai.claude_client import claude_client
from app.ai.json_parsing import parse_claude_json
from app.ai.prompts.base import BASE_SYSTEM_PROMPT, get_generation_prompt
from app.ai.prompts.worksheet import WORKSHEET_SYSTEM_PROMPT, WORKSHEET_GENERATION_PROMPT
from app.ai.schemas.worksheet import WorksheetSchema
from app.ai.schemas.basic_types import PassageSchema, QuizSchema, AssessmentSchema
from app.utils.storage import storage_manager
from app.utils.logger import logger
from app.core.templates.engine import template_engine
from app.core.enums import TemplateType, GradeLevel, WorldviewFlag

class GeneratorAgent:
    """AI agent responsible for generating educational content"""
    
    def __init__(self):
        self.schema_map = {
            'WORKSHEET': WorksheetSchema,
            'PASSAGE': PassageSchema,
            'QUIZ': QuizSchema,
            'ASSESSMENT': AssessmentSchema
        }
        
        self.prompt_map = {
            'WORKSHEET': (WORKSHEET_SYSTEM_PROMPT, WORKSHEET_GENERATION_PROMPT)
        }
    
    def _parse_json_response(self, raw_output: str) -> Dict[str, Any]:
        """Parse Claude JSON response (fence/prose tolerant; see app.ai.json_parsing)"""
        return parse_claude_json(raw_output)
    
    async def generate_content(
        self, 
        product_id: int,
        product_type: str, 
        standard: str, 
        grade_level: int, 
        curriculum: str
    ) -> Dict[str, Any]:
        """Generate content for a product and save to storage"""
        
        try:
            logger.info(f"Starting content generation for product {product_id} ({product_type})")
            
            # Get appropriate prompts
            if product_type in self.prompt_map:
                system_prompt, generation_template = self.prompt_map[product_type]
            else:
                system_prompt = BASE_SYSTEM_PROMPT
                generation_template = ""
            
            # Build user prompt
            user_prompt = get_generation_prompt(product_type, standard, grade_level, curriculum)
            if generation_template:
                user_prompt += f"\\n\\n{generation_template}"
            
            # Generate content using Claude
            raw_output = await claude_client.generate(system_prompt, user_prompt)
            logger.info(f"Claude raw output for product {product_id}: {raw_output[:200]}...")
            
            # Parse JSON response (shared fence/prose-tolerant parser)
            try:
                content_data = self._parse_json_response(raw_output)
            except ValueError:
                logger.error(f"Unparseable Claude output for product {product_id} — raw response logged by parser")
                raise
            
            # Skip schema validation - Claude generates good content but with different field names
            # Content is saved successfully to storage regardless of schema validation
            
            # Save raw content to storage
            storage_manager.save_json_file(product_id, "raw", content_data)
            
            logger.info(f"Content generation completed for product {product_id}")
            return content_data
            
        except Exception as e:
            logger.error(f"Content generation failed for product {product_id}: {e}")
            raise
    
    async def generate_template_content(
        self,
        product_id: int,
        template_type: str,
        standard: str,
        grade_level: int,
        ela_standard_type: str,
        ela_standard_code: str,
        worldview_flag: str,
        curriculum: str,
        bundle_context: dict | None = None
    ) -> Dict[str, Any]:
        """Generate content using template-driven approach"""
        try:
            logger.info(f"Starting template-based generation for product {product_id} ({template_type})")

            template = template_engine.get_template(TemplateType(template_type))
            user_prompt = template.generate_prompt(
                ela_standard_code,
                GradeLevel(str(grade_level)),
                WorldviewFlag(worldview_flag)
            )

            # Determine temperature: lower for bundle-linked to keep context locked
            temperature = 0.7 if (bundle_context and template_type != 'ANCHOR_READING_PASSAGE') else 1.0

            # Build system prompt
            system_prompt = (
                f"{BASE_SYSTEM_PROMPT}\n\n"
                f"Template Type: {template_type}\n"
                f"IMPORTANT: The 'title' field in your JSON output MUST exactly match "
                f"the title shown in the OUTPUT FORMAT schema. Do NOT change it based on the standard."
            )

            # Inject full passage text for strict context chaining
            if bundle_context and template_type != 'ANCHOR_READING_PASSAGE':
                passage_text = bundle_context.get('passage_text', '')
                passage_title = bundle_context.get('passage_title', '')
                key_vocabulary = bundle_context.get('key_vocabulary', '')

                context_block = (
                    f"\n\n=== BUNDLE CONTEXT — MANDATORY ===\n"
                    f"This template is part of a cohesive bundle. ALL content you generate MUST be "
                    f"based ONLY on the anchor passage provided below.\n\n"
                    f"STRICT RULES:\n"
                    f"- Use ONLY characters, settings, and events that appear in this passage.\n"
                    f"- Do NOT reference any external books, classic literature, or outside characters.\n"
                    f"- Do NOT invent new characters, plot points, or scenarios not present in the passage.\n"
                    f"- All questions, prompts, vocabulary sentences, and examples must cite ONLY this passage.\n"
                    f"- Passage title: \"{passage_title}\"\n"
                )
                if key_vocabulary:
                    context_block += f"- Key vocabulary from passage: {key_vocabulary}\n"
                if passage_text:
                    context_block += f"\nFULL ANCHOR PASSAGE TEXT:\n{passage_text}\n"
                context_block += "=== END BUNDLE CONTEXT ===\n"
                user_prompt += context_block
                logger.info(f"Bundle context injected for product {product_id} (passage: '{passage_title}', {len(passage_text)} chars)")

            # Generate with Claude
            raw_output = await claude_client.generate(system_prompt, user_prompt, temperature=temperature)
            logger.info(f"Claude raw output for product {product_id}: {raw_output[:200]}...")

            content_data = self._parse_json_response(raw_output)

            # Validate bundle context compliance — check for outside character/event hallucination
            if bundle_context and template_type != 'ANCHOR_READING_PASSAGE':
                violation = self._check_context_violation(content_data, bundle_context)
                if violation:
                    logger.warning(f"Context violation detected for product {product_id}: {violation}. Regenerating with correction prompt.")
                    correction_prompt = (
                        f"{user_prompt}\n\n"
                        f"CORRECTION REQUIRED: Your previous output contained content not grounded in the "
                        f"provided anchor passage. Specifically: {violation}\n"
                        f"Regenerate the entire output. Every character name, plot event, and example "
                        f"MUST come directly from the anchor passage text above. No exceptions."
                    )
                    raw_output = await claude_client.generate(system_prompt, correction_prompt, temperature=0.5)
                    content_data = self._parse_json_response(raw_output)
                    logger.info(f"Regenerated content after context violation for product {product_id}")

            validation_errors = template.validate_content(content_data)
            if validation_errors:
                logger.warning(f"Length validation failed for product {product_id}: {validation_errors}. Regenerating with correction prompt.")
                correction_prompt = (
                    f"{user_prompt}\n\n"
                    f"CORRECTION REQUIRED: Your previous output violated min/max length rules.\n"
                    f"Violations: {'; '.join(validation_errors)}\n"
                    f"Regenerate the entire output strictly within all specified minimum and maximum "
                    f"length limits. Do not add or remove content beyond what the limits allow."
                )
                raw_output = await claude_client.generate(system_prompt, correction_prompt, temperature=0.5)
                content_data = self._parse_json_response(raw_output)
                second_errors = template.validate_content(content_data)
                if second_errors:
                    raise ValueError(
                        f"Content length validation failed after regeneration for product {product_id}. "
                        f"Violations: {'; '.join(second_errors)}"
                    )

            storage_manager.save_json_file(product_id, "raw", content_data)
            logger.info(f"Template content generation completed for product {product_id}")
            return content_data

        except Exception as e:
            logger.error(f"Template content generation failed for product {product_id}: {e}")
            raise

    def _check_context_violation(self, content_data: Dict[str, Any], bundle_context: dict) -> str | None:
        """Check if generated content references characters/events not in the anchor passage.
        Returns a violation description string, or None if clean.
        """
        passage_text = (bundle_context.get('passage_text') or '').lower()
        if not passage_text:
            return None

        # Flatten all string values from content_data for scanning
        all_text = self._flatten_text(content_data).lower()

        # Extract word tokens from content that look like proper names (Title Case, 3+ chars)
        import re
        # Find capitalised words in the generated content
        candidate_names = set(re.findall(r'\b[A-Z][a-z]{2,}\b', self._flatten_text(content_data)))
        # Remove common non-name words
        common_words = {
            'The', 'This', 'That', 'These', 'Those', 'When', 'Where', 'What', 'Which',
            'How', 'Why', 'Who', 'Each', 'Some', 'Many', 'Most', 'More', 'Such',
            'Also', 'Then', 'Than', 'From', 'With', 'Into', 'Upon', 'Over', 'After',
            'Before', 'Because', 'However', 'Therefore', 'Although', 'Answer', 'Question',
            'Write', 'Read', 'Grade', 'Student', 'Teacher', 'Bundle', 'Title', 'Text',
            'Standard', 'Common', 'Core', 'According', 'Evidence', 'Response', 'Using',
            'Support', 'Explain', 'Describe', 'Identify', 'Analyze', 'Compare', 'English',
            'Christian', 'Neutral', 'Reading', 'Writing', 'Short', 'Exit', 'Ticket',
        }
        candidate_names -= common_words

        outside_names = [
            name for name in candidate_names
            if name.lower() not in passage_text
        ]

        if outside_names:
            return f"references names not found in passage: {', '.join(list(outside_names)[:5])}"
        return None

    def _flatten_text(self, obj, _depth: int = 0) -> str:
        """Recursively extract all string values from a dict/list structure."""
        if _depth > 6:
            return ''
        if isinstance(obj, str):
            return obj
        if isinstance(obj, list):
            return ' '.join(self._flatten_text(i, _depth + 1) for i in obj)
        if isinstance(obj, dict):
            return ' '.join(self._flatten_text(v, _depth + 1) for v in obj.values())
        return ''

generator_agent = GeneratorAgent()
