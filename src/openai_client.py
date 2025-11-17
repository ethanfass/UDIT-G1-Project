"""OpenAI API client wrapper for questionnaire analysis."""

from openai import OpenAI
from src.config import Config


class QuestionnaireAnalyzer:
    """Wrapper for OpenAI API interactions."""
    
    def __init__(self):
        """Initialize the OpenAI client."""
        Config.validate()
        self.client = OpenAI(api_key=Config.OPENAI_API_KEY)
        self.model = Config.OPENAI_MODEL
    
    def analyze_questionnaire(self, questionnaire_text: str, rubric: str) -> dict:
        """
        Analyze a questionnaire against a rubric using OpenAI.
        
        Args:
            questionnaire_text: The questionnaire content to analyze
            rubric: The rubric/standards to evaluate against
            
        Returns:
            Dictionary containing the analysis results
        """
        prompt = f"""
        Analyze the following questionnaire against the provided rubric.
        
        RUBRIC/STANDARDS:
        {rubric}
        
        QUESTIONNAIRE:
        {questionnaire_text}
        
        Please provide:
        1. A score based on the rubric
        2. Key findings
        3. Areas of concern
        4. Recommendations for improvement
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
            return {
                "status": "success",
                "analysis": response.choices[0].message.content
            }
        except Exception as e:
            return {
                "status": "error",
                "error": str(e)
            }
    
    def convert_to_template(self, questionnaire_text: str, template: str) -> dict:
        """
        Convert a questionnaire to match a standard template.
        
        Args:
            questionnaire_text: The original questionnaire
            template: The template format to convert to
            
        Returns:
            Dictionary containing the converted questionnaire
        """
        prompt = f"""
        Please convert the following questionnaire to match this template format:
        
        TEMPLATE:
        {template}
        
        ORIGINAL QUESTIONNAIRE:
        {questionnaire_text}
        
        Return the questionnaire in the template format, maintaining all relevant information.
        """
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "user", "content": prompt}
                ],
                temperature=0.5,
                max_tokens=2000
            )
            
            return {
                "status": "success",
                "converted": response.choices[0].message.content
            }
        except Exception as e:
            return {
                "status": "error",
                "error": str(e)
            }


if __name__ == "__main__":
    print("OpenAI client module loaded")
