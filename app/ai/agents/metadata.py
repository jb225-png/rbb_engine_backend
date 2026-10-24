import json
from typing import Dict, Any
from app.ai.claude_client import claude_client
from app.ai.json_parsing import parse_claude_json
from app.ai.schemas.metadata import MetadataSchema
from app.utils.storage import storage_manager
from app.utils.logger import logger

class MetadataAgent:
    """AI agent responsible for generating product metadata"""
    
    METADATA_SYSTEM_PROMPT = """You are an expert educational content metadata generator.

RESPONSIBILITIES:
- Generate compelling titles and descriptions
- Create relevant tags and SEO keywords
- Suggest appropriate pricing
- Classify difficulty and skill levels
- Identify learning outcomes and usage contexts

METADATA REQUIREMENTS:
- Title: Engaging, descriptive, SEO-friendly
- Description: Clear value proposition, key features
- Tags: Relevant, searchable, curriculum-aligned
- Keywords: High-traffic educational search terms
- Price: Market-appropriate for content quality and scope

PRICING GUIDELINES:
- Simple worksheets: $0.99-$2.99
- Complex activities: $3.99-$7.99
- Assessment packages: $8.99-$15.99
- Comprehensive units: $16.99-$29.99

OUTPUT: Valid JSON only, following metadata schema exactly."""
    
    def _parse_json_response(self, raw_output: str) -> Dict[str, Any]:
        """Parse Claude JSON response (fence/prose tolerant; see app.ai.json_parsing)"""
        return parse_claude_json(raw_output)
    
    async def generate_metadata(
        self,
        product_id: int,
        product_type: str,
        content: Dict[str, Any],
        standard: str,
        grade_level: int,
        curriculum: str
    ) -> Dict[str, Any]:
        """Generate metadata for content and save to storage"""
        
        try:
            logger.info(f"Starting metadata generation for product {product_id}")
            
            # Build metadata generation prompt
            content_summary = self._extract_content_summary(content, product_type)
            
            user_prompt = f"""Generate comprehensive metadata for this {product_type.lower()}:

CONTENT SUMMARY:
{content_summary}

EDUCATIONAL CONTEXT:
- Standard: {standard}
- Grade Level: {grade_level}
- Curriculum: {curriculum}

GENERATE:
- Compelling title (10-100 chars)
- Marketing description (50-300 chars)
- Relevant tags (3-10 items)
- SEO keywords (5-15 items)
- Appropriate price ($0.99-$99.99)
- Difficulty classification
- Learning outcomes
- Usage recommendations

Ensure metadata is market-ready and educationally accurate."""
            
            # Generate metadata using Claude
            raw_output = await claude_client.generate(self.METADATA_SYSTEM_PROMPT, user_prompt)
            
            # Parse JSON response
            metadata_data = self._parse_json_response(raw_output)
            
            # Save metadata to storage
            storage_manager.save_json_file(product_id, "metadata", metadata_data)
            
            logger.info(f"Metadata generation completed for product {product_id}")
            return metadata_data
            
        except Exception as e:
            logger.error(f"Metadata generation failed for product {product_id}: {e}")
            # Generate and save fallback metadata
            fallback_metadata = self._generate_fallback_metadata(product_type, grade_level, standard)
            storage_manager.save_json_file(product_id, "metadata", fallback_metadata)
            return fallback_metadata
    
    async def generate_template_metadata(
        self,
        product_id: int,
        template_type: str,
        content: Dict[str, Any],
        standard: str,
        grade_level: int,
        ela_standard_code: str,
        worldview_flag: str
    ) -> Dict[str, Any]:
        """Generate SEO metadata for template content"""
        try:
            logger.info(f"Starting template metadata generation for product {product_id}")
            
            # Build metadata generation prompt
            content_summary = self._extract_content_summary(content, template_type)
            
            user_prompt = f"""Generate SEO metadata for {template_type}:

CONTENT SUMMARY:
{content_summary}

EDUCATIONAL CONTEXT:
- Standard: {ela_standard_code}
- Grade Level: {grade_level}
- Worldview: {worldview_flag}

GENERATE JSON with:
- seo_title (60 chars, keyword-rich)
- seo_description (160 chars, compelling)
- internal_linking_block (HTML snippet with related content links)
- social_snippets (JSON array of 3 social media posts)

Ensure metadata is market-ready and educationally accurate."""
            
            # Generate metadata using Claude
            raw_output = await claude_client.generate(self.METADATA_SYSTEM_PROMPT, user_prompt)
            
            # Parse JSON response
            metadata_data = self._parse_json_response(raw_output)
            
            # Save metadata to storage
            storage_manager.save_json_file(product_id, "metadata", metadata_data)
            
            logger.info(f"Template metadata generation completed for product {product_id}")
            return metadata_data
            
        except Exception as e:
            logger.error(f"Template metadata generation failed for product {product_id}: {e}")
            # Generate and save fallback metadata
            fallback_metadata = {
                "seo_title": f"Grade {grade_level} {template_type.replace('_', ' ').title()} - {ela_standard_code}",
                "seo_description": f"Christian-facing ELA {template_type.lower()} for grade {grade_level} aligned with {ela_standard_code} standard.",
                "internal_linking_block": f"<p>Related resources for <strong>{ela_standard_code}</strong></p>",
                "social_snippets": json.dumps([
                    f"New Grade {grade_level} ELA resource: {template_type.replace('_', ' ')}",
                    f"Master {ela_standard_code} with our {template_type.replace('_', ' ').lower()}",
                    f"Christian-aligned ELA content for middle school educators"
                ])
            }
            storage_manager.save_json_file(product_id, "metadata", fallback_metadata)
            return fallback_metadata
    
    def _extract_content_summary(self, content: Dict[str, Any], product_type: str) -> str:
        """Extract key information from content for metadata generation"""
        summary_parts = []
        
        if 'title' in content:
            summary_parts.append(f"Title: {content['title']}")
        
        if 'learning_objectives' in content:
            objectives = content['learning_objectives'][:2]  # First 2 objectives
            summary_parts.append(f"Objectives: {', '.join(objectives)}")
        
        if 'questions' in content:
            summary_parts.append(f"Questions: {len(content['questions'])} items")
        
        if 'estimated_time' in content:
            summary_parts.append(f"Duration: {content['estimated_time']} minutes")
        
        return " | ".join(summary_parts) if summary_parts else f"{product_type} content"
    
    def _generate_fallback_metadata(self, product_type: str, grade_level: int, standard: str) -> Dict[str, Any]:
        """Generate basic fallback metadata when AI generation fails"""
        return {
            "title": f"Grade {grade_level} {product_type.title()} - {standard[:50]}",
            "description": f"Educational {product_type.lower()} aligned with {standard} for grade {grade_level} students. Includes comprehensive activities and answer key.",
            "tags": [product_type.lower(), f"grade-{grade_level}", "education", "curriculum", "worksheet"],
            "seo_keywords": [f"grade {grade_level}", product_type.lower(), "education", "curriculum", "learning"],
            "suggested_price": 2.99,
            "difficulty_level": "intermediate",
            "estimated_duration": 45,
            "learning_outcomes": [
                f"Students will understand key concepts from {standard}",
                f"Students will apply grade {grade_level} appropriate skills"
            ],
            "subject_area": "General Education",
            "topic_focus": standard[:100],
            "skill_level": "developing",
            "classroom_ready": True,
            "homework_suitable": True,
            "assessment_type": "formative"
        }

metadata_agent = MetadataAgent()