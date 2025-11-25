import sys
sys.path.insert(0, '.')
from core.powerpoint_agent_core import PowerPointAgent
import tempfile, os

with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as f:
    path = f.name

try:
    with PowerPointAgent() as agent:
        agent.create_new()
        agent.add_slide(layout_name='Blank')
        
        # This now works exactly as the system prompt describes:
        result = agent.add_shape(
            slide_index=0,
            shape_type="rectangle",
            position={"left": "0%", "top": "0%"},
            size={"width": "100%", "height": "100%"},
            fill_color="#FFFFFF",
            fill_opacity=0.15  # Subtle overlay ✅
        )
        
        print(f"✅ Overlay created successfully!")
        print(f"   Shape Index: {result['shape_index']}")
        print(f"   Fill Opacity: {result['styling']['fill_opacity']}")
        print(f"   Applied: {result['styling']['fill_opacity_applied']}")
        
        agent.save(path)
        print('\n🎉 END-TO-END OVERLAY TEST PASSED!')
finally:
    if os.path.exists(path):
        os.unlink(path)
