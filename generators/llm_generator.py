"""
LLM-powered presentation structure generator using LangChain and OpenAI.
"""
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import HumanMessage

from schemas import PresentationInput

SYSTEM_PROMPT = """You are a senior presentation strategist creating professional PowerPoint presentations.

## OUTPUT FORMAT

Generate a JSON presentation with this structure:
- title: Presentation title
- subtitle: Optional tagline
- author: Optional author name
- subject: Optional subject description
- slides: Array of slides (count will be specified in the request)

## SLIDE TYPES TO USE (mix them!)

**Text slide**: title + content (paragraph) or bullet_points (list)
**Data slide**: title + table (with headers, rows, style)
**Chart slide**: title + chart (column/line/pie with categories and series)
**Visual slide**: title + image (path + position) + optional bullet_points
**Combined slide**: title + content + bullet_points (intro text + key points)

## MANDATORY ELEMENTS (MUST INCLUDE)

Every presentation MUST contain:
- At least 2 slides with IMAGES (use layout: image_right or image_left)
- At least 1 TABLE with meaningful data
- At least 1 CHART showing trends or comparisons
- Mix of content paragraphs and bullet_points

Image path should be descriptive keywords for image search, like:
- "ancient mystical ceremony ritual"
- "modern technology digital innovation"
- "professional business team collaboration"
- "nature landscape scenic view"

## COMPLETE EXAMPLE

{{
  "title": "Strategic Analysis 2024",
  "subtitle": "Market Opportunities and Challenges",
  "author": "Strategy Team",
  "subject": "Annual strategic review",
  "slides": [
    {{
      "title": "Executive Summary",
      "content": "This presentation analyzes our market position and identifies key growth opportunities for the coming year.",
      "layout": "content_only"
    }},
    {{
      "title": "Key Findings",
      "bullet_points": [
        "Market share increased by 15% in Q3",
        "Customer satisfaction at all-time high of 92%",
        "Three new market segments identified",
        "Cost reduction of 8% achieved"
      ],
      "layout": "content_only"
    }},
    {{
      "title": "Performance Metrics",
      "table": {{
        "headers": ["Metric", "2023", "2024", "Growth"],
        "rows": [
          ["Revenue", "$2.1M", "$2.8M", "+33%"],
          ["Customers", "1,200", "1,850", "+54%"],
          ["NPS Score", "72", "85", "+18%"],
          ["Market Share", "12%", "18%", "+50%"]
        ],
        "style": "header_colored"
      }},
      "layout": null
    }},
    {{
      "title": "Revenue Trend",
      "chart": {{
        "type": "column",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
          {{"name": "2023", "data": [450, 520, 580, 550]}},
          {{"name": "2024", "data": [520, 680, 750, 850]}}
        ]
      }},
      "layout": null
    }},
    {{
      "title": "Market Expansion",
      "image": {{
        "path": "global business expansion world map",
        "position": "right"
      }},
      "bullet_points": [
        "European market entry planned for Q2",
        "Partnership with local distributors",
        "Initial investment of $500K allocated"
      ],
      "layout": "image_right"
    }},
    {{
      "title": "Innovation & Technology",
      "image": {{
        "path": "modern technology digital transformation office",
        "position": "left"
      }},
      "bullet_points": [
        "AI-powered analytics platform launched",
        "Cloud infrastructure migration complete",
        "Mobile app reaching 50K downloads"
      ],
      "layout": "image_left"
    }},
    {{
      "title": "Next Steps",
      "content": "Based on our analysis, we recommend the following strategic priorities for the coming quarter.",
      "bullet_points": [
        "Accelerate European expansion",
        "Invest in customer success team",
        "Launch new product line by Q3",
        "Continue cost optimization program"
      ],
      "layout": "content_only"
    }}
  ]
}}

## QUALITY GUIDELINES
- Use SPECIFIC data relevant to the topic (not generic placeholders)
- Tables should have 3-5 meaningful rows
- Charts should show realistic trends with 2+ data points
- Bullet points should be concrete and actionable
- Mix different slide types for visual variety
- Every slide must have substantial content
"""


def create_presentation_agent(api_key: str):
    """Creates a LangChain agent for structured presentation generation."""
    llm = ChatOpenAI(
        model="gpt-4o-2024-08-06",
        temperature=0.6,
        api_key=api_key
    )

    formatting_strategy = llm.with_structured_output(PresentationInput, method="json_schema")

    prompt = ChatPromptTemplate.from_messages([
        ("system", SYSTEM_PROMPT),
        MessagesPlaceholder(variable_name="messages")
    ])

    return prompt | formatting_strategy


def generate_presentation_structure(
    topic: str, 
    api_key: str, 
    slide_count: int | None = None, 
    language: str = "en"
) -> PresentationInput:
    """
    Generate a presentation structure from a topic using LLM.
    
    Args:
        topic: The presentation subject
        api_key: OpenAI API key
        slide_count: Target number of slides (5-15)
        language: Output language ('en' or 'fr')
    
    Returns:
        PresentationInput: Validated presentation structure
    """
    agent = create_presentation_agent(api_key)
    
    count_clause = f" Target {slide_count} slides." if slide_count else ""
    lang_clause = f" Write the content in {'French' if language == 'fr' else 'English'}."
    
    response = agent.invoke({
        "messages": [HumanMessage(content=f"Create a presentation about: {topic}.{count_clause}{lang_clause}")]
    })
    
    return response
