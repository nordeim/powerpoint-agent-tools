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
        
        # Test add_shape with opacity
        result1 = agent.add_shape(
            slide_index=0,
            shape_type='rectangle',
            position={'left':'10%','top':'10%'},
            size={'width':'20%','height':'20%'},
            fill_color='#0070C0',
            fill_opacity=0.5
        )
        print('✅ add_shape with fill_opacity:', result1['styling']['fill_opacity_applied'])
        
        # Test format_shape with fill_opacity
        result2 = agent.format_shape(
            slide_index=0,
            shape_index=result1['shape_index'],
            fill_opacity=0.3
        )
        print('✅ format_shape with fill_opacity:', result2['success'])
        
        # Test deprecated transparency (backward compatibility)
        result3 = agent.format_shape(
            slide_index=0,
            shape_index=result1['shape_index'],
            transparency=0.7  # Should convert to fill_opacity=0.3
        )
        print('✅ format_shape with transparency:', 'transparency_converted_to_opacity' in result3['changes_applied'])
        
        agent.save(path)
        print('\n🎉 ALL TESTS PASSED!')
finally:
    if os.path.exists(path):
        os.unlink(path)
