import json
from typing import Dict, Any
from app.ai.claude_client import claude_client
from app.ai.json_parsing import parse_claude_json
from app.ai.prompts.base import get_qc_prompt
from app.ai.schemas.qc import QCSchema
from app.utils.storage import storage_manager
from app.utils.logger import logger

class QCAgent:
    """AI agent responsible for quality control evaluation"""
    
    QC_SYSTEM_PROMPT = """You are an expert educational content quality evaluator.

EVALUATION CRITERIA:
1. STRUCTURE: Organization, completeness, format consistency
2. ALIGNMENT: Match with educational standard and curriculum
3. CLARITY: Clear instructions, understandable language
4. DIFFICULTY: Appropriate for grade level, progressive challenge
5. INCLUSIVITY: Culturally sensitive, accessible to diverse learners
6. ACCURACY: Factually correct, pedagogically sound

SCORING SCALE:
- 90-100: Excellent, ready for immediate use
- 75-89: Good, minor improvements needed
- 60-74: Adequate, needs fixes before use
- 40-59: Poor, significant issues present
- 0-39: Unacceptable, major revision required

VERDICT GUIDELINES:
- PASS: Score >= 75, minor or no issues
- NEEDS_FIX: Score 50-74, fixable issues identified
- FAIL: Score < 50, major problems requiring regeneration

OUTPUT: Valid JSON only, following QC schema exactly."""
    
    def _parse_json_response(self, raw_output: str) -> Dict[str, Any]:
        """Parse Claude JSON response (fence/prose tolerant; see app.ai.json_parsing)"""
        return parse_claude_json(raw_output)
    
    async def evaluate_content(
        self,
        product_id: int,
        product_type: str,
        content: Dict[str, Any],
        standard: str,
        grade_level: int
    ) -> Dict[str, Any]:
        """Evaluate content quality and save QC results"""
        
        try:
            logger.info(f"Starting QC evaluation for product {product_id}")
            
            # Convert content to string for evaluation
            content_str = json.dumps(content, indent=2)
            
            # Build QC prompt
            user_prompt = get_qc_prompt(product_type, content_str, standard, grade_level)
            
            # Get QC evaluation from Claude
            raw_output = await claude_client.generate(self.QC_SYSTEM_PROMPT, user_prompt)
            
            # Parse JSON response
            qc_data = self._parse_json_response(raw_output)
            
            # Save QC results to storage
            storage_manager.save_json_file(product_id, "qc", qc_data)
            
            logger.info(f"QC evaluation completed for product {product_id}: {qc_data['verdict']} (score: {qc_data['score']})")
            return qc_data
            
        except Exception as e:
            logger.error(f"QC evaluation failed for product {product_id}: {e}")
            # Return failure QC result
            fallback_qc = {
                "verdict": "FAIL",
                "score": 0,
                "issues": [{
                    "category": "structure",
                    "severity": "critical",
                    "description": f"QC process failed: {str(e)}",
                    "suggestion": "Retry content generation"
                }],
                "structure_score": 0,
                "alignment_score": 0,
                "clarity_score": 0,
                "difficulty_score": 0,
                "inclusivity_score": 0,
                "accuracy_score": 0,
                "strengths": ["None identified"],
                "recommendations": ["Retry generation process"]
            }
            storage_manager.save_json_file(product_id, "qc", fallback_qc)
            return fallback_qc
    
    async def evaluate_template_content(
        self,
        product_id: int,
        template_type: str,
        content: Dict[str, Any],
        standard: str,
        grade_level: int,
        worldview_flag: str
    ) -> Dict[str, Any]:
        """Evaluate template content with worldview validation"""
        try:
            logger.info(f"Starting template QC evaluation for product {product_id}")
            
            # Convert content to string for evaluation
            content_str = json.dumps(content, indent=2)
            
            # Build enhanced QC prompt
            user_prompt = f"""Evaluate this {template_type} content:

CONTENT:
{content_str}

EVALUATION CRITERIA:
- Template structure compliance
- Grade {grade_level} appropriateness
- Standard alignment: {standard}
- Worldview: {worldview_flag}
"""
            
            if worldview_flag == "CHRISTIAN":
                user_prompt += """
CHRISTIAN CONTENT VALIDATION:
- Values-aligned but not preachy
- Natural integration of themes
- Academic tone maintained
"""
            
            user_prompt += "\n\nProvide evaluation as JSON with verdict (PASS/NEEDS_FIX/FAIL), score (0-100), and specific issues list."
            
            # Get QC evaluation from Claude
            raw_output = await claude_client.generate(self.QC_SYSTEM_PROMPT, user_prompt)
            
            # Parse JSON response
            qc_data = self._parse_json_response(raw_output)
            
            # Save QC results to storage
            storage_manager.save_json_file(product_id, "qc", qc_data)
            
            logger.info(f"Template QC evaluation completed for product {product_id}: {qc_data.get('verdict', 'UNKNOWN')} (score: {qc_data.get('score', 0)})")
            return qc_data
            
        except Exception as e:
            logger.error(f"Template QC evaluation failed for product {product_id}: {e}")
            # Return failure QC result
            fallback_qc = {
                "verdict": "FAIL",
                "score": 0,
                "issues": [{
                    "category": "structure",
                    "severity": "critical",
                    "description": f"QC process failed: {str(e)}",
                    "suggestion": "Retry content generation"
                }],
                "structure_score": 0,
                "alignment_score": 0,
                "clarity_score": 0,
                "difficulty_score": 0,
                "inclusivity_score": 0,
                "accuracy_score": 0,
                "strengths": ["None identified"],
                "recommendations": ["Retry generation process"]
            }
            storage_manager.save_json_file(product_id, "qc", fallback_qc)
            return fallback_qc

qc_agent = QCAgent()