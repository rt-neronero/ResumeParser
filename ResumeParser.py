import docx
import re
import json

class ResumeParser:
    def __init__(self, resume_path):
        self.resume_path = resume_path
        self.resume = {
            'personal_info': {},
            'employment_details': [],
            'education_details': [],
            'skils': []
        }
        self._parse_resume()

    def _parse_resume(self):
        # Open the DOCX file and extract the text
        doc = docx.Document(self.resume_path)
        full_text = '\n'.join([para.text for para in doc.paragraphs])

        print("Full Text of Resume:") # debugger
        print("----------------------") # debugger
        print(full_text)
        print()

        # Call specific functions to parse sections
        self._parse_personal_info(full_text)
        self._get_experience_section(full_text)
        self._get_education_section(full_text)
        self._get_skills_section(full_text)

    def _parse_personal_info(self, text):
        # Use regex for better name detection (full name including middle names)
        name_pattern = re.compile(r"([A-Za-z]+(?: (?:[A-Za-z]+\b|[A-Za-z]\.)){0,4} [A-Za-z]+)")  # Handles first, middle, and last names
        name_match = name_pattern.search(text)
        if name_match:
            self.resume['personal_info']['Full Name'] = name_match.group(1).strip()

        # Regex for email extraction (to capture multiple emails)
        email_pattern = re.compile(r"([\w.-]+@[\w.-]+)")
        email_matches = email_pattern.findall(text)
        if email_matches:
            self.resume['personal_info']['Email'] = list(set(email_matches))  # Remove duplicates, store emails in a list

        # Regex for phone number extraction (including country code and '09' starting numbers)
        phone_pattern = re.compile(r"(\+?\d{1,2}\s?)?(\(?\d{2,4}\)?[\s-]?)?\(?\d{3,4}\)?[\s-]?\d{3,4}[\s-]?\d{3,4}")
        phone_match = phone_pattern.search(text)
        if phone_match:
            self.resume['personal_info']['Phone Number'] = phone_match.group(0).strip()

        # Regex for links (social media, portfolio, etc.)
        link_pattern = re.compile(r"(?<![\w.@])((https?://)?(?:www\.)?[a-zA-Z0-9-]+\.[a-z]{2,6}(?:/[^\s]*)?)")
        links = link_pattern.findall(text)
        self.resume['personal_info']['Links'] = [link[0] for link in links]

    def _get_experience_section(self, text):
        # Find Experience section
        experience_pattern = re.compile(r"(^|\n)(experience|work experience|background work experience)(\n|:|\s+)", re.IGNORECASE)
        
        # Find the start of the Experience section
        start_index_match = experience_pattern.search(text)
        if start_index_match:
            start_index = start_index_match.end()  # Move past the "Experience" section header
            
            # Find the next section header to determine the end of the experience section
            end_index = len(text)
            next_section_headers = ["Education", "Skills", "Certifications", "Languages", "Personal"]
            for header in next_section_headers:
                header_index = text.lower().find(header.lower(), start_index)
                if header_index != -1 and header_index < end_index:
                    end_index = header_index

            # Extract the experience text
            experience_text = text[start_index:end_index].strip()
            
            if experience_text:
                print("Experience Details Extracted:") #debugger
                print(experience_text) #debugger
                self._parse_experience_details(experience_text)
            else:
                print("No content found in the Experience section.")
        else:
            print("Experience Section not found.")
    
    def _parse_experience_details(self, experience_text):
        pass
    def _get_education_section(self, text):
        # Look for the Education section header
        education_pattern = re.compile(r"(^|\n)(education|academic background)(\n|:|\s+)", re.IGNORECASE)
        
        # Find the start of the Education section
        start_index_match = education_pattern.search(text)
        if start_index_match:
            start_index = start_index_match.end()  # Move past the "Education" section header
            
            # Look for the next section headers to determine the end of the education section
            end_index = len(text)
            next_section_headers = ["Skills", "Certifications", "Languages", "Experience", "Personal"]
            for header in next_section_headers:
                header_index = text.lower().find(header.lower(), start_index)
                if header_index != -1 and header_index < end_index:
                    end_index = header_index
            
            # Extract the education text
            education_text = text[start_index:end_index].strip()
            
            if education_text:
                print("\nEducation Section Extracted:")
                print(education_text)
                self._parse_education_details(education_text)
            else:
                print("No content found in the Education section.")
        else:
            print("Education section not found.")

    def _parse_education_details(self, text):
        # Regex for education
        pass
    def _get_skills_section(self, text):
        # Look for the Skills section header
        skills_pattern = re.compile(r"(^|\n)(skills|core competencies)(\n|:)", re.IGNORECASE)
        
        # Find the start of the Skills section
        start_index_match = skills_pattern.search(text)
        
        if start_index_match:
            start_index = start_index_match.end()  # Move past the "Skills" section header
            
            # Look for the next section headers to determine the end of the skills section
            end_index = len(text)
            next_section_headers = ["Experience", "Education", "Certifications", "Languages", "Personal"]
            for header in next_section_headers:
                header_index = text.lower().find(header.lower(), start_index)
                if header_index != -1 and header_index < end_index:
                    end_index = header_index
            
            # Extract the skills text
            skills_text = text[start_index:end_index].strip()
            
            if skills_text:
                print("\nSkills Section Extracted:")
                print(skills_text)
                self._parse_skills_details(skills_text)
            else:
                print("No content found in the Skills section.") #debugger

        else:
            print("Skills section not found.") #debugger

    def _parse_skills_details(self, text):
        pass
    def get_parsed_data(self):
        return json.dumps(self.resume, indent=4)

# Sample usage
resume_parser = ResumeParser("Logarta_CV.docx")
parsed_data = resume_parser.get_parsed_data()
print(parsed_data)
