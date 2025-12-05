"""
Museum Story Generator - Multi-LLM Support
Generates narrative stories for museum objects using various AI models
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
import time
import os
from typing import Optional

# LLM Configuration
LLM_CONFIGS = {
    'groq': {
        'env_key': 'TOKEN',
        'model': 'llama-3.3-70b-versatile',
        'library': 'groq'
    },
    'openai': {
        'env_key': 'TOKEN',
        'model': 'gpt-4o',
        'library': 'openai'
    },
    'gemini': {
        'env_key': 'TOKEN',
        'model': 'gemini-2.0-flash-exp',
        'library': 'google-generativeai'
    },
    'anthropic': {
        'env_key': 'TOKEN',
        'model': 'claude-sonnet-4-5-20250929',
        'library': 'anthropic'
    },
    'phi4': {
        'env_key': 'TOKEN',
        'model': 'Phi-4',
        'library': 'langchain',
        'base_url': 'https://models.inference.ai.azure.com'
    }
}

STORY_PROMPT = """You are a knowledgeable and passionate museum curator with expertise in archaeology and cultural heritage.

Below is structured information about an archaeological object from the MUDEC museum's collection. Your task is to narrate a vivid, emotionally engaging story of the object.

Please speak in a respectful, accessible tone — balancing storytelling with historical context.

GUIDELINES:
- Write for the ear, not the eye: use flowing, conversational language that sounds natural when spoken
- Use natural transitions between ideas ("Moving forward in time...", "Centuries later...", "Now picture...")
- Provide clear spatial descriptions ("roughly the size of a small table...", "smooth as polished stone")
- IMPORTANT: The final text must be between 170 and 270 words
- The text must be purely narrative. Do NOT include lists, bullet points, tables, or metadata

OBJECT INFORMATION (structured data from Neo4j knowledge graph):
{}

NARRATIVE STRUCTURE FOR AUDIO:
1. Opening Hook (25–35 words): Begin with an evocative scene or intriguing detail that immediately engages the listener
2. Physical Description (20–30 words): Help listeners visualize the object using size comparisons, material descriptions, and tactile qualities
3. Historical Relevance (50–70 words): Inform visitors about the historical background of the piece and the cultural context it comes from
4. Historical Journey (60–80 words): Tell the story chronologically with clear time markers and smooth transitions between eras
5. Present-Day Connection (25–30 words): Conclude by linking the object's journey to today, emphasizing its relevance to visitors
"""


class StoryGenerator:
    def __init__(self, provider='groq', model=None, api_key=None):
        self.provider = provider.lower()
        self.config = LLM_CONFIGS.get(self.provider)
        
        if not self.config:
            raise ValueError(f"Unknown provider: {provider}. Choose from: {list(LLM_CONFIGS.keys())}")
        
        self.api_key = api_key or os.getenv(self.config['env_key'])
        if not self.api_key:
            raise ValueError(f"API key not found. Set TOKEN env variable or pass api_key parameter")
        
        self.model = model or self.config['model']
        self.client = self._init_client()
    
    def _init_client(self):
        if self.config['library'] == 'groq':
            from groq import Groq
            return Groq(api_key=self.api_key)
        
        elif self.config['library'] == 'openai':
            from openai import OpenAI
            return OpenAI(api_key=self.api_key)
        
        elif self.config['library'] == 'langchain':
            from langchain_openai import ChatOpenAI
            return ChatOpenAI(
                model=self.model,
                api_key=self.api_key,
                base_url=self.config.get('base_url'),
                temperature=0.7
            )
        
        elif self.config['library'] == 'google-generativeai':
            import google.generativeai as genai
            genai.configure(api_key=self.api_key)
            return genai.GenerativeModel(self.model)
        
        elif self.config['library'] == 'anthropic':
            from anthropic import Anthropic
            return Anthropic(api_key=self.api_key)
        
        else:
            raise ValueError(f"Unknown library: {self.config['library']}")
    
    def generate(self, description, retries=3):
        if not description or description.strip() == '' or description == 'No information available':
            return "Insufficient information to generate a narrative for this object."
        
        prompt = STORY_PROMPT.format(description)
        
        for attempt in range(retries):
            try:
                if self.config['library'] == 'groq':
                    response = self.client.chat.completions.create(
                        model=self.model,
                        messages=[{"role": "user", "content": prompt}],
                        max_tokens=500,
                        temperature=0.7
                    )
                    return response.choices[0].message.content.strip()
                
                elif self.config['library'] == 'openai':
                    response = self.client.chat.completions.create(
                        model=self.model,
                        messages=[{"role": "user", "content": prompt}],
                        max_tokens=500,
                        temperature=0.7
                    )
                    return response.choices[0].message.content.strip()
                
                elif self.config['library'] == 'langchain':
                    from langchain.schema import HumanMessage, SystemMessage
                    messages = [
                        SystemMessage(content="You are an expert museum curator specializing in archaeological storytelling."),
                        HumanMessage(content=prompt)
                    ]
                    response = self.client.invoke(messages)
                    return response.content.strip()
                
                elif self.config['library'] == 'google-generativeai':
                    response = self.client.generate_content(prompt)
                    return response.text.strip()
                
                elif self.config['library'] == 'anthropic':
                    message = self.client.messages.create(
                        model=self.model,
                        max_tokens=1024,
                        temperature=0.7,
                        system="You are an expert museum curator specializing in archaeological storytelling.",
                        messages=[{"role": "user", "content": prompt}]
                    )
                    return message.content[0].text.strip()
                
            except Exception as e:
                error = str(e)
                print(f"    ⚠️  Attempt {attempt + 1}/{retries} failed: {error[:100]}")
                
                if "rate_limit" in error.lower() or "429" in error:
                    print(f"    ⏳ Rate limited, waiting 60s...")
                    time.sleep(60)
                elif "timeout" in error.lower():
                    time.sleep(10)
                elif attempt < retries - 1:
                    time.sleep(5)
                else:
                    return f"Error after {retries} attempts: {error[:200]}"
        
        return "Failed to generate story after multiple attempts."


def process_catalog(input_file, output_file, provider='groq', model=None, 
                   api_key=None, save_interval=10):
    
    print(f"📖 Reading {input_file}...")
    print(f"🤖 Provider: {provider}")
    if model:
        print(f"🎯 Model: {model}")
    print()
    
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return
    
    print(f"✅ Found {len(df)} objects\n")
    
    generator = StoryGenerator(provider=provider, model=model, api_key=api_key)
    
    if 'Narrative Story' not in df.columns:
        df['Narrative Story'] = ''
    
    existing = df['Narrative Story'].notna().sum()
    if existing > 0:
        print(f"📊 {existing} stories already exist\n")
    
    total = len(df)
    success = 0
    errors = 0
    start = time.time()
    
    for idx, row in df.iterrows():
        if pd.notna(row['Narrative Story']) and row['Narrative Story'] != '':
            success += 1
            print(f"⏭️  {idx + 1}/{total}: {row['Object ID']} - Skip")
            continue
        
        obj_id = row['Object ID']
        desc = row['Description']
        
        print(f"✍️  {idx + 1}/{total}: {obj_id}")
        
        try:
            story = generator.generate(desc)
            df.at[idx, 'Narrative Story'] = story
            
            if "Error" in story or "Failed" in story or "Insufficient" in story:
                errors += 1
                print(f"    ⚠️  Generation issue")
            else:
                success += 1
                wc = len(story.split())
                print(f"    ✅ Success ({wc} words)")
        
        except Exception as e:
            df.at[idx, 'Narrative Story'] = f"Error: {str(e)}"
            errors += 1
            print(f"    ❌ {e}")
        
        if (idx + 1) % save_interval == 0:
            elapsed = time.time() - start
            avg = elapsed / (idx + 1 - existing) if (idx + 1 - existing) > 0 else 0
            remaining = (total - idx - 1) * avg if avg > 0 else 0
            
            print(f"  💾 Checkpoint ({idx + 1}/{total})")
            print(f"  ⏱️  {avg:.1f}s/obj - {remaining/60:.1f}m left")
            
            try:
                df.to_excel(output_file, index=False)
                print(f"  ✅ Saved - Success: {success}, Errors: {errors}\n")
            except Exception as e:
                print(f"  ⚠️  Save error: {e}\n")
        
        time.sleep(1)
    
    print(f"\n💾 Final save to {output_file}...")
    try:
        df.to_excel(output_file, index=False, engine='openpyxl')
        
        wb = load_workbook(output_file)
        ws = wb.active
        
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 100
        ws.column_dimensions['D'].width = 80
        
        for cell in ws[1]:
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=3).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            ws.cell(row=r, column=4).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            ws.cell(row=r, column=1).alignment = Alignment(horizontal='center', vertical='top')
            ws.cell(row=r, column=2).alignment = Alignment(horizontal='left', vertical='top')
            ws.row_dimensions[r].height = 300
        
        wb.save(output_file)
        
        total_time = time.time() - start
        
        print(f"\n✅ Complete!")
        print(f"📊 Summary:")
        print(f"   Total: {total}")
        print(f"   Success: {success}")
        print(f"   Errors: {errors}")
        print(f"   Success rate: {(success/total)*100:.1f}%")
        print(f"   Time: {total_time/60:.1f}m")
        if (total - existing) > 0:
            print(f"   Avg: {total_time/(total-existing):.1f}s/obj")
        
    except Exception as e:
        print(f"❌ Save error: {e}")


if __name__ == "__main__":
    process_catalog(
        input_file="catalog_100_objects.xlsx",
        output_file="catalog_with_stories.xlsx",
        provider='groq',  # Options: 'groq', 'openai', 'gemini', 'anthropic', 'phi4'
        model=None,  # Optional: override default model
        save_interval=10
    )
