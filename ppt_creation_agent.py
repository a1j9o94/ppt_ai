import json
from openai import OpenAI
from typing import List, Dict, Any
import os
from dotenv import load_dotenv
import logging
from ppt_helpers import create_ppt, add_slide as add_slide_to_ppt, delete_ppt
from models import PresentationModel, SlideModel, Section, LayoutType, PresentationRequest

# Load environment variables
load_dotenv()

# Set up logging with the same format as ppt_api.py
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Define the available functions that match your ppt_api.py routes
def create_presentation(name: str, title: str, slides: List[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Create a new PowerPoint presentation with the given title."""
    logger.debug(f"Creating presentation with title: {title}")
    try:
        file_path = create_ppt(name, title, slides)
        return {"file_path": file_path, "title": title}
    except Exception as e:
        logger.error(f"Failed to create presentation: {str(e)}")
        return {"error": str(e)}

def add_slide(ppt_name: str, slide_title: str, columns: int = None, rows: int = None, layout: str = None, sections: List[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Add a new slide to the presentation."""
    logger.debug(f"Adding slide to presentation {ppt_name} with layout: {layout}")
    logger.debug(f"Slide sections: {sections}")
    try:
        sections = sections or []
        if layout == 'COLUMN':
            add_slide_to_ppt(ppt_name, slide_title, columns=len(sections), sections=sections)
        elif layout == 'ROW':
            add_slide_to_ppt(ppt_name, slide_title, rows=len(sections), sections=sections)
        else:
            add_slide_to_ppt(ppt_name, slide_title)
        return {"message": f"Slide '{slide_title}' added to '{ppt_name}' successfully"}
    except Exception as e:
        logger.error(f"Failed to add slide: {str(e)}")
        return {"error": str(e)}

def save_presentation(name: str) -> Dict[str, Any]:
    """Save the presentation and return the file path."""
    logger.debug(f"Saving presentation: {name}")
    try:
        file_path = f"output/{name}.pptx"
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Presentation '{name}' does not exist.")
        # Additional save logic can be added here if needed.
        return {"file_path": file_path}
    except Exception as e:
        logger.error(f"Failed to save presentation: {str(e)}")
        return {"error": str(e)}

def delete_presentation(name: str) -> Dict[str, Any]:
    """Delete the presentation."""
    logger.debug(f"Deleting presentation: {name}")
    try:
        delete_ppt(name)
        return {"message": f"Presentation '{name}' deleted successfully"}
    except Exception as e:
        logger.error(f"Failed to delete presentation: {str(e)}")
        return {"error": str(e)}

# Define the function descriptions for the AI model
FUNCTION_DESCRIPTIONS = [
    {
        "name": "create_presentation",
        "description": "Create a new PowerPoint presentation",
        "parameters": PresentationModel.model_json_schema()
    },
    {
        "name": "add_slide",
        "description": "Add a new slide to the presentation",
        "parameters": {
            "type": "object",
            "properties": {
                **SlideModel.model_json_schema()["properties"]
            },
            "required": ["ppt_name", "slide_title"]
        }
    },
    {
        "name": "save_presentation",
        "description": "Save the presentation and get the file path",
        "parameters": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "The name of the presentation to save"
                }
            },
            "required": ["name"]
        }
    },
    {
        "name": "delete_presentation",
        "description": "Delete the presentation",
        "parameters": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "The name of the presentation to delete"
                }
            },
            "required": ["name"]
        }
    }
]

def create_presentation_from_prompt(prompt: str, client=None) -> Dict[str, Any]:
    """
    Create a PowerPoint presentation based on the given prompt using the OpenAI API.
    """

    if client is None:
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    logger.debug(f"Creating presentation from prompt: {prompt}")

    # Update example slides to use the new models
    example_slide_rows = SlideModel(
        slide_title="Sized Row Layout",
        layout=LayoutType.ROW,
        rows=3,
        sections=[
            Section(header="Row 1", content=["Content 1"], size=33),
            Section(header="Row 2", content=["Content 2", "Content 3", "Content 4"], size=33),
            Section(header="Row 3", content=["Content 2", "Content 3", "Content 4"], size=33)
        ]
    ).model_dump()

    example_slide_columns = SlideModel(
        name="test_file",
        slide_title="Sized Column Layout",
        layout=LayoutType.COLUMN,
        columns=3,
        sections=[
            Section(header="Wide Column", content=["Content 1"], size=50),
            Section(header="Narrow Column", content=["Content 2", "Content 3", "Content 4"], size=25),
            Section(header="Narrow Column", content=["Content 2", "Content 3", "Content 4"], size=25)
        ]
    ).model_dump()

    user_prompt = f'''
        You are an AI assistant tasked with helping consultants outline and structure presentations for their cases. Your job is to create a well-organized and detailed PowerPoint presentation based on the given subject. Your objective is not just to outline high-level bullet points but to also provide detailed descriptions or narratives for each section, explaining the insights or conclusions to be drawn from the content.

        Instructions:
        The presentation is about: <presentation_subject> {prompt} </presentation_subject>

        You have access to the following tools to create the presentation:

        create_presentation: Creates a new PowerPoint presentation with a short and concise title.
        add_slide: Adds a new slide to the presentation with a specified title, layout, and content.
        save_presentation: Saves the presentation and returns the file path.
        delete_presentation: Deletes the presentation if needed.

        Here are a few examples of how to structure the presentation input:
        {example_slide_rows}
        {example_slide_columns}

        Analyze the topic thoroughly to generate a detailed slide outline. Typically include:

        A title slide
        Content slides (main body of the presentation, breaking down each point in detail)
        A conclusion or summary slide
        For each slide in your outline, focus on the key insight or message that should be conveyed in that section. Use the add_slide function to create a new slide. When creating a slide:

        Provide a clear and concise slide_title
        Choose an appropriate layout (e.g., 'COLUMN' for multi-column layouts, 'ROW' for rows)
        Be thoughtful about what portion of the page should go to each topic? E.g. is it 66% a graph with 33% of commentary
        Structure the content with detailed descriptions of each point
        Go beyond bullet points and provide a brief narrative explaining why each point is important or relevant to the overall goal of the presentation.
        When adding content to slides:

        Use clear, concise language but add brief descriptions to explain each bullet point
        Break down complex ideas into easy-to-understand concepts
        Use placeholders like "[INSERT DATA]" or "[ADD SPECIFIC EXAMPLE]" for missing information
        Ensure a logical flow from one slide to the next, with each slide building on the previous one.
        For data-heavy sections, create slides with appropriate layouts to accommodate charts, graphs, or tables. Use placeholders like "[INSERT CHART: Sales Growth]" and provide a brief description of what the chart or data is supposed to demonstrate.

        After creating all slides, save the completed presentation using the save_presentation function.

        If you need to start over at any point, use the delete_presentation function to remove the current presentation before creating a new one.

        The following is an example of a good detailed presentation outline for a presentation about pricing strategies for IBM:

        1.	Meeting objectives
        a.	Recap TLS program findings and pricing’s role in future transformation
        b.	Discuss pricing insights informed by historical analysis, competitor, and customer perspectives 
        c.	Share high-level roadmap for pricing over the next 3-6 months 
        d.	Align on pricing path forward 
        2.	Recap of TLS strategy and last meeting – slide per each bullet below
        a.	Strategy roadmap: Pricing is a key element of building an integrated Systems and TLS strategy (TLS roadmap page with box over integrated strategy bullet)
        b.	Pricing framework: In our last meeting, we discussed diving deeper into key pricing topics across Systems and TLS; today we will focus on three main areas (visual page with price setting, price realization and delegations, pricing roadmap and icons)
        3.	Price setting (at POS and renewal) – slide per each bullet below
        a.	Pricing overview: IBM's current approach to price setting involves input from many stakeholders, outside of Expert Care (two panel slide with overview of price setting process for Expert Care and non-Expert Care; include current share of HW with EC vs. not EC across served equipment) 
        b.	Comparison vs. competitors: Competitors take a more uniform approach to price setting for services at POS and renewal (Comparative case study between IBM and Dell / HPE, including price setting as % of HW, P&L management, YoY increases – to be developed)
        c.	Customer expectations as % of HW: Most customers expect services to be priced at 5-20% of the hardware price annually; customers typically refresh between 3-6 years (Phase 1 survey slide)
        d.	TPM pricing models: TPMs typically use a fixed annual contract model to develop service prices (Case study on Park Place, Service Express, Evernex price setting for services, including initial price, YoY changes, term lengths, etc. – to be developed)
        e.	Customer expectations for YoY price increase: Customers expect ~4% annual increases to maintenance services costs; APAC expects slightly higher increases than Americas and EMEA (Phase 1 survey slide)
        f.	Price as driver of switching to TPM: When faced with price increase that would put IT over-budget, up to XX% of customers indicate they would switch to a TPM (Phase 1 survey slide)
        4.	Price realization and delegations – slide per each bullet below
        a.	Historic price realization: Historically, price realization has ranged from XX-XX% across geos and varies at POS vs. renewal (Services price realization analysis from Dusan // whatever we can get from their team on historic average discounts, ideally cut by geography and POS / renewal)
        b.	Current delegation approach: Current TLS delegations include four levels but process varies by geography (Current overview of TLS delegations, differences by geo – qualitative) 
        c.	HPE case study: HPE uses a joint “deal desk” to ensure margins are maintained across Systems and services (Case study on HPE and their joint "deal desk" that manages discounts for entire deal – to be developed)
        d.	Dell case study: Dell uses XYZ approach to discounting and does Y to maintain margins (Case study on Dell and how they run price delegations (potential to combine with above depending on level of differentiation) – to be developed)
        e.	Price sensitivity: Price sensitivity charts indicate that different cohorts of customers may be driving differences in sensitivity (from Phase 1 survey) 
        f.	Expected discount levels: Most customers prefer 3- to 4-year contracts & expect discount of ~10-15% compared to 1-year contracts (from Phase 1 survey)
        5.	Roadmap and next steps – slide per each bullet below
        a.	Roadmap: Revising our pricing strategy can be accomplished with buy-in across Systems and TLS (3-6 month roadmap of what we would do within Systems / TLS, who would need to be involved and key actions – potentially by month or quarter)
        b.	Near-term actions: Our progress to date and where we should go next (two panel summary slide – LHS: What actions we have taken so far – what is the impact, RHS: What actions have we yet to take?) 
        c.	Next steps (bulleted list of next steps / follow up items)


        Error Handling:
        Handle any errors that may occur during the process. If a function call fails, explain the error and suggest a solution or alternative approach.
        Output Structure:
        Once the presentation is completed, provide a detailed summary of the structure, including:

        The total number of slides
        A detailed description of each slide's content and message (not just bullet points, but also the narrative or point to be conveyed)
        Any areas where additional data or input is needed from the consulting team
        Present the final output in the following format:

        Presentation Summary:

        [Include your summary here]
        Slide Outline:

        [List each slide with its title and a description of its main points and narrative]
        Next Steps:

        [Provide suggestions for what the consulting team should do next to finalize the presentation]
    '''
    
    # Start a conversation with the AI
    messages = [{"role": "user", "content": user_prompt}]
    
    presentation_data = {"status": "in_progress"}
    
    while presentation_data["status"] == "in_progress" or presentation_data["status"] == "error":
        logger.debug("Making API call to OpenAI")
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=messages,
                functions=FUNCTION_DESCRIPTIONS,
                function_call="auto"
            )
            
            response_message = response.choices[0].message
            finish_reason = response.choices[0].finish_reason
            logger.debug(f"Assistant's finish_reason: {finish_reason}")
            
            # If the model wants to call a function
            if response_message.function_call:
                function_name = response_message.function_call.name
                function_args = json.loads(response_message.function_call.arguments)
                logger.debug(f"AI requested function call: {function_name}")
                logger.debug(f"Function arguments: {function_args}")
                
                function_response = None
                # Call the appropriate function
                try:
                    if "create_presentation" in str(function_name):
                        function_response = create_presentation(**function_args)
                        presentation_data = {"status": "in_progress", **function_response}
                    elif "add_slide" in str(function_name):
                        function_response = add_slide(**function_args)
                        presentation_data = {"status": "in_progress", **function_response}
                    elif "save_presentation" in str(function_name):
                        function_response = save_presentation(**function_args)
                        presentation_data = {"status": "completed", **function_response}
                    elif "delete_presentation" in str(function_name):
                        function_response = delete_presentation(**function_args)
                        presentation_data = {"status": "completed", **function_response}
                    
                    logger.debug(f"Function response: {function_response}")
                except Exception as e:
                    logger.error(f"Error executing function {function_name}: {str(e)}", exc_info=True)
                    function_response = {"error": str(e)}
                    presentation_data = {"status": "error", "error": str(e)}
                
                # Add the function response to the conversation
                messages.append({
                    "role": "function",
                    "name": function_name,
                    "content": json.dumps(function_response)
                })
            else:
                # Assistant's response does not include a function call
                if finish_reason == "stop":
                    # Assistant has finished its final reply
                    logger.debug("Assistant indicated conversation is complete.")
                    presentation_data["status"] = "completed"
                else:
                    # Assistant may have more to say; continue the loop
                    logger.debug("Assistant may provide more responses; continuing the loop.")
            
            # Add the assistant's response to the messages
            messages.append(response_message)
            
        except Exception as e:
            logger.error(f"Error in OpenAI API call: {str(e)}", exc_info=True)
            raise
    
    logger.debug("Presentation creation completed")
    return presentation_data

if __name__ == "__main__":
    # Example usage
    prompt = "Create a 5 page presentation about how FMCG companies could use generative AI to improve their business"
    try:
        logger.info(f"Starting presentation creation with prompt: {prompt}")
        result = create_presentation_from_prompt(prompt)
        logger.info(f"Presentation created successfully: {result}")
    except Exception as e:
        logger.error(f"Failed to create presentation: {str(e)}", exc_info=True)
