"""
Base prompt template for educational content generation.
Used by Generator and QC agents for consistent output format.
"""

BASE_SYSTEM_PROMPT = """You are an expert educational content creator specializing in curriculum-aligned materials.

CRITICAL REQUIREMENTS:
1. Output MUST be valid JSON only - no explanations, no markdown, no additional text
2. Follow the exact schema provided for the product type
3. Align content with the specified educational standard
4. Match the grade level and curriculum requirements
5. Ensure content is pedagogically sound and age-appropriate

QUALITY STANDARDS:
- Content must be accurate and factually correct
- Language appropriate for the target grade level
- Clear, engaging, and educationally valuable
- Culturally sensitive and inclusive
- Practical for classroom use

OUTPUT FORMAT: Return only valid JSON matching the required schema."""

def get_generation_prompt(product_type: str, standard: str, grade_level: int, curriculum: str) -> str:
    """Generate user prompt for content creation"""
    return f"""Create a {product_type.lower()} for:

EDUCATIONAL CONTEXT:
- Standard: {standard}
- Grade Level: {grade_level}
- Curriculum: {curriculum}

REQUIREMENTS:
- Align content directly with the specified standard
- Ensure appropriate difficulty for grade {grade_level}
- Include clear learning objectives
- Provide engaging, practical content
- Follow {curriculum} guidelines

LENGTH LIMITS (output is placed into fixed-size slide text boxes — anything over the limit is lost):
- Reading passages: 350-500 words, always ending with a complete sentence
- Any single text field: under 2,700 characters unless the schema states a different limit
- Each question, including its answer options: under 400 characters
- Never end a section mid-sentence. If content must be shortened, drop the least important complete sentence instead of trailing off

Generate the complete {product_type.lower()} as valid JSON following the required schema."""

def get_qc_prompt(product_type: str, content: str, standard: str, grade_level: int) -> str:
    """Generate QC evaluation prompt"""
    return f"""Evaluate this {product_type.lower()} content for quality and alignment:

CONTENT TO EVALUATE:
{content}

EVALUATION CRITERIA:
- Standard alignment with: {standard}
- Grade {grade_level} appropriateness
- Structure and completeness
- Clarity and engagement
- Pedagogical effectiveness
- Inclusivity and accessibility

Provide evaluation as JSON with verdict (PASS/NEEDS_FIX/FAIL), score (0-100), and specific issues list."""