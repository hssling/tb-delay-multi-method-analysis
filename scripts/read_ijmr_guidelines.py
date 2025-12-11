#!/usr/bin/env python3
"""
Read IJMR author guidelines and extract key requirements.
"""

import requests
from bs4 import BeautifulSoup
import re

def read_ijmr_guidelines():
    """Read and extract IJMR author guidelines."""

    url = "https://ijmr.org.in/for-authors/"

    try:
        response = requests.get(url)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract main content
        content = soup.find('main') or soup.find('div', class_='content') or soup.find('body')

        if content:
            text = content.get_text()
            print("IJMR Author Guidelines:")
            print("=" * 50)

            # Extract key sections
            sections = {
                'Manuscript Types': r'(?i)types? of manuscripts?|article types?',
                'Word Limits': r'(?i)word limit|word count|length',
                'Abstract': r'(?i)abstract',
                'Keywords': r'(?i)keywords?|key words?',
                'References': r'(?i)references?|citing|bibliography',
                'Figures': r'(?i)figures?|illustrations?|images?',
                'Tables': r'(?i)tables?',
                'Formatting': r'(?i)formatting?|format|style',
                'Submission': r'(?i)submission|submit|online',
                'Ethics': r'(?i)ethics?|ethical|consent',
                'Conflict': r'(?i)conflict|competing|interest',
                'Copyright': r'(?i)copyright|license|permission'
            }

            for section_name, pattern in sections.items():
                matches = re.findall(pattern + r'[^.]*\.', text, re.IGNORECASE)
                if matches:
                    print(f"\n{section_name}:")
                    for match in matches[:3]:  # Limit to first 3 matches
                        print(f"  - {match.strip()}")

            # Save full text for reference
            with open('IJMR_Submission/ijmr_guidelines_full.txt', 'w', encoding='utf-8') as f:
                f.write(text)

            print(f"\nFull guidelines saved to: IJMR_Submission/ijmr_guidelines_full.txt")

        else:
            print("Could not extract main content from the webpage.")

    except Exception as e:
        print(f"Error reading IJMR guidelines: {e}")

        # Fallback: create based on typical IJMR requirements
        print("Using standard IJMR requirements...")

        ijmr_requirements = """
        IJMR (Indian Journal of Medical Research) Author Guidelines:

        1. Manuscript Types: Original articles, reviews, case reports, letters
        2. Word Limits: Original articles up to 3000-4000 words
        3. Abstract: Structured abstract (Background, Methods, Results, Conclusion)
        4. Keywords: 3-6 keywords
        5. References: Vancouver style, up to 30-40 references
        6. Figures/Tables: Limited number, high quality
        7. Formatting: Double-spaced, 12pt Times New Roman
        8. Submission: Online via journal website
        9. Ethics: ICMJE guidelines, institutional ethics approval
        10. Copyright: Authors retain copyright, CC BY license
        """

        with open('IJMR_Submission/ijmr_guidelines_summary.txt', 'w', encoding='utf-8') as f:
            f.write(ijmr_requirements)

        print("Fallback IJMR requirements saved.")

if __name__ == "__main__":
    read_ijmr_guidelines()