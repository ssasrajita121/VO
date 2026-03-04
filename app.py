import streamlit as st
import requests
import json
import os
import tempfile
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import google.generativeai as genai
from dotenv import load_dotenv
import time

# Load environment
load_dotenv()

# Page config
st.set_page_config(page_title="EduBridge Voice-Over Generator", page_icon="🎙️", layout="wide")

# CSS
st.markdown("""<style>
.main-header{font-size:2.5rem;font-weight:bold;color:#2D3E6D;text-align:center;margin-bottom:0.5rem}
.stButton>button{background-color:#2D3E6D;color:white;font-size:1.1rem;padding:0.75rem 2rem;border-radius:0.5rem}
</style>""", unsafe_allow_html=True)

# API Keys
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY", "")
SPEAKATOO_API_KEY = "9GS8VO5RM10077052a1d1da894b9a19cd31909e0cZHB3Nci2X"

# Speakatoo Configuration - CORRECT API v1
SPEAKATOO_CONFIG = {
    "api_url": "https://www.speakatoo.com/api/v1/voiceapi",
    "api_key": SPEAKATOO_API_KEY,
    "username": "richa@edubridgeindia.in",
    "password": "Siddh@0410",
    "voice_id": "BFUw72Nl589b0c29fbff4cf7c8c97d2d8bd0818afFpy9aNxI1",  # Neerja Neural
    "engine": "neural",
    "format": "mp3"
}

def extract_slide_content(slide):
    """Extract text content from a slide"""
    content = []
    if slide.shapes.title:
        content.append(f"Title: {slide.shapes.title.text}")
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            if shape != slide.shapes.title:
                content.append(shape.text.strip())
    return "\n".join(content) if content else "Slide content"

def generate_voice_script(slide_content, allocated_seconds, slide_number, total_slides):
    """Generate voice-over script using Gemini with strict limits"""
    try:
        genai.configure(api_key=GOOGLE_API_KEY)
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        # Calculate word limit (2.5 words per second)
        max_words = int(allocated_seconds * 2.5)
        
        prompt = f"""Create a voice-over script for PowerPoint slide {slide_number} of {total_slides}.

STRICT REQUIREMENTS:
- Maximum {max_words} words (approximately {allocated_seconds} seconds)
- NO markdown formatting (**bold**, *italic*, `code`)
- Plain English text only
- No special characters or symbols
- Natural, conversational tone
- Professional and clear

Content: {slide_content}

Return ONLY the narration script in plain text. No formatting, no extra words."""

        response = model.generate_content(prompt)
        script = response.text.strip()
        
        # Strip any markdown that slipped through
        import re
        script = re.sub(r'\*\*(.+?)\*\*', r'\1', script)  # Remove **bold**
        script = re.sub(r'\*(.+?)\*', r'\1', script)      # Remove *italic*
        script = re.sub(r'`(.+?)`', r'\1', script)        # Remove `code`
        script = script.replace('**', '').replace('*', '').replace('`', '')
        
        # Enforce word limit strictly
        words = script.split()
        if len(words) > max_words:
            script = ' '.join(words[:max_words])
            st.warning(f"Slide {slide_number}: Trimmed to {max_words} words")
        
        return script
        
    except Exception as e:
        st.error(f"Gemini Error: {str(e)}")
        # Fallback with word limit
        fallback = f"Slide {slide_number}. {slide_content}"
        return ' '.join(fallback.split()[:max_words])

def generate_audio_speakatoo(text, filename="VoiceOver"):
    """Generate audio using Speakatoo API v1"""
    try:
        headers = {
            "X-API-KEY": SPEAKATOO_CONFIG["api_key"],
            "Content-Type": "application/json"
        }
        
        payload = {
            "username": SPEAKATOO_CONFIG["username"],
            "password": SPEAKATOO_CONFIG["password"],
            "tts_title": filename,
            "ssml_mode": "0",
            "tts_engine": SPEAKATOO_CONFIG["engine"],
            "tts_format": SPEAKATOO_CONFIG["format"],
            "tts_text": text,
            "tts_resource_ids": SPEAKATOO_CONFIG["voice_id"],
            "synthesize_type": "save"
        }
        
        response = requests.post(
            SPEAKATOO_CONFIG["api_url"],
            json=payload,
            headers=headers,
            timeout=60
        )
        
        if response.status_code == 200:
            result = response.json()
            
            if result.get("result") or result.get("status"):
                audio_url = result.get("tts_uri")
                
                if audio_url:
                    return audio_url
                else:
                    st.error(f"No audio URL: {result}")
                    return None
            else:
                st.error(f"API Error: {result.get('message', result)}")
                return None
        else:
            st.error(f"HTTP {response.status_code}: {response.text[:200]}")
            return None
            
    except Exception as e:
        st.error(f"Exception: {str(e)}")
        return None

def add_audio_to_slide(slide, audio_url):
    """Download and embed audio into slide with speaker icon"""
    try:
        import requests
        import tempfile
        
        # Download the MP3 file
        audio_response = requests.get(audio_url, timeout=30)
        
        if audio_response.status_code == 200:
            # Save to temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as tmp_audio:
                tmp_audio.write(audio_response.content)
                tmp_audio_path = tmp_audio.name
            
            # Insert audio into slide
            # Position for speaker icon (bottom-right)
            left = Inches(8.5)
            top = Inches(4.8)
            
            # Add audio to slide - creates speaker icon automatically
            # By default, audio plays on click in PowerPoint
            movie = slide.shapes.add_movie(
                tmp_audio_path,
                left, top,
                width=Inches(0.5),
                height=Inches(0.5),
                poster_frame_image=None,  # Use default speaker icon
                mime_type='audio/mp3'
            )
            
            # Audio plays on click by default - no need to set action
            
            # Clean up temp file
            try:
                os.remove(tmp_audio_path)
            except:
                pass
            
            return True
        else:
            st.error(f"❌ Failed to download audio: {audio_response.status_code}")
            return False
            
    except Exception as e:
        st.error(f"Error embedding audio: {str(e)}")
        return False

def process_presentation(uploaded_file, target_duration_minutes=60):
    """Process presentation and add voice-overs"""
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_path = tmp_file.name
    
    try:
        prs = Presentation(tmp_path)
        total_slides = len(prs.slides)
        
        st.info(f"📊 {total_slides} slides | ⏱️ {target_duration_minutes} min target")
        
        total_seconds = target_duration_minutes * 60
        seconds_per_slide = total_seconds / total_slides
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        success_count = 0
        total_chars = 0  # Track total characters
        
        for idx, slide in enumerate(prs.slides, 1):
            status_text.text(f"Processing slide {idx}/{total_slides}...")
            
            slide_content = extract_slide_content(slide)
            
            voice_script = generate_voice_script(
                slide_content, 
                int(seconds_per_slide), 
                idx, 
                total_slides
            )
            
            char_count = len(voice_script)
            word_count = len(voice_script.split())
            total_chars += char_count  # Add to total
            
            with st.expander(f"📝 Slide {idx} Script: {word_count} words, {char_count} chars"):
                st.write(voice_script)
                st.caption(f"⏱️ Target: {int(seconds_per_slide)}s | Words: {word_count} | Characters: {char_count}")
            
            audio_url = generate_audio_speakatoo(voice_script, f"Slide_{idx}")
            
            if audio_url:
                if add_audio_to_slide(slide, audio_url):
                    success_count += 1
                    st.success(f"✅ Slide {idx}: Audio added")
                else:
                    st.warning(f"⚠️ Slide {idx}: Audio generated but icon failed")
            else:
                st.error(f"❌ Slide {idx}: Audio generation failed")
            
            progress_bar.progress(idx / total_slides)
            time.sleep(0.5)
        
        output_path = tmp_path.replace('.pptx', '_voiceover.pptx')
        prs.save(output_path)
        
        status_text.text("✅ Complete!")
        progress_bar.progress(1.0)
        
        # Show character usage summary
        st.success(f"""
        📊 **Character Usage Summary:**
        - Total characters: **{total_chars:,}**
        - Average per slide: **{total_chars // total_slides:,}**
        - Estimated Speakatoo cost: Based on your plan
        """)
        
        return output_path, success_count, total_slides, total_chars
    
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None, 0, 0, 0  # Added total_chars
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

# UI
st.markdown('<div class="main-header">🎙️ EduBridge Voice-Over Generator</div>', unsafe_allow_html=True)
st.markdown('<div style="text-align:center;color:#2D3E6D;margin-bottom:2rem">Add AI Narration with Neerja Voice</div>', unsafe_allow_html=True)

with st.expander("📖 How It Works"):
    st.markdown("""
    1. **Upload** PowerPoint (.pptx)
    2. **Set duration** (default: 60 minutes)
    3. **Generate** - AI creates scripts and audio
    4. **Download** presentation with 🔊 icons
    5. **Click 🔊** during presentation to play audio
    
    **Voice:** Neerja (Female, Indian English, Neural)
    """)

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader("Upload PowerPoint", type=['pptx'])
    target_duration = st.slider("Duration (minutes)", 10, 120, 60, 5)

with col2:
    st.info(f"""
    **Configuration:**
    - Voice: Neerja 🎤
    - Language: English (India)
    - Engine: Neural AI
    - Format: MP3
    """)
    
    if GOOGLE_API_KEY:
        st.success("✅ Gemini ready")
    else:
        st.error("⚠️ Gemini key missing")
    
    st.success("✅ Speakatoo ready")

st.markdown("---")

if uploaded_file:
    if st.button("🎙️ Generate Voice-Overs", use_container_width=True):
        if not GOOGLE_API_KEY:
            st.error("⚠️ Set GOOGLE_API_KEY in .env")
        else:
            output_path, success, total, total_chars = process_presentation(uploaded_file, target_duration)
            
            if output_path and os.path.exists(output_path):
                st.success(f"🎉 Added voice-overs to {success}/{total} slides!")
                st.info(f"📊 Total characters used: **{total_chars:,}**")
                
                with open(output_path, 'rb') as f:
                    pptx_data = f.read()
                
                st.download_button(
                    "📥 Download Presentation",
                    pptx_data,
                    f"{uploaded_file.name.replace('.pptx', '')}_voiceover.pptx",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
                
                os.remove(output_path)
                
                st.info("✅ Open in PowerPoint and click 🔊 icons to play audio")
else:
    st.info("👆 Upload a PowerPoint to begin")

st.markdown("---")
st.markdown('<div style="text-align:center;color:#2D3E6D">🎓 EduBridge | Powered by Gemini + Speakatoo (Neerja)</div>', unsafe_allow_html=True)
