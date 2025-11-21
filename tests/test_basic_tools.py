#!/usr/bin/env python3
"""
PowerPoint Agent Tools - Basic Integration Tests
Test the first 5 P0 tools

Run with: pytest test_basic_tools.py -v
Or: python test_basic_tools.py
"""

import pytest
import json
import subprocess
import sys
import tempfile
import shutil
from pathlib import Path


class TestBasicTools:
    """Test basic PowerPoint tool functionality."""
    
    @pytest.fixture
    def temp_dir(self):
        """Create temporary directory for test files."""
        temp_path = tempfile.mkdtemp()
        yield Path(temp_path)
        shutil.rmtree(temp_path, ignore_errors=True)
    
    @pytest.fixture
    def tools_dir(self):
        """Get tools directory path."""
        return Path(__file__).parent / 'tools'
    
    def run_tool(self, tool_name: str, args: dict, tools_dir: Path) -> dict:
        """Run tool and return parsed JSON response."""
        cmd = [sys.executable, str(tools_dir / tool_name), '--json']
        
        for key, value in args.items():
            if isinstance(value, bool):
                if value:
                    cmd.append(f'--{key}')
            elif isinstance(value, dict):
                # JSON argument
                cmd.extend([f'--{key}', json.dumps(value)])
            else:
                cmd.extend([f'--{key}', str(value)])
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        try:
            data = json.loads(result.stdout)
            return {
                'returncode': result.returncode,
                'data': data,
                'stderr': result.stderr
            }
        except json.JSONDecodeError:
            return {
                'returncode': result.returncode,
                'data': {},
                'stderr': result.stderr,
                'stdout': result.stdout
            }
    
    def test_create_new_basic(self, tools_dir, temp_dir):
        """Test creating new presentation."""
        output = temp_dir / 'test.pptx'
        
        result = self.run_tool('ppt_create_new.py', {
            'output': output,
            'slides': 3
        }, tools_dir)
        
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_create_new.py\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        assert result['data']['status'] == 'success'
        assert output.exists()
        assert result['data']['slides_created'] == 3
    
    def test_create_new_with_layout(self, tools_dir, temp_dir):
        """Test creating with specific layout."""
        output = temp_dir / 'layout_test.pptx'
        
        result = self.run_tool('ppt_create_new.py', {
            'output': output,
            'slides': 5,
            'layout': 'Title and Content'
        }, tools_dir)
        
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_create_new.py\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        assert output.exists()
        assert 'available_layouts' in result['data']
    
    def test_add_slide(self, tools_dir, temp_dir):
        """Test adding slide to existing presentation."""
        # First create presentation
        filepath = temp_dir / 'add_slide_test.pptx'
        self.run_tool('ppt_create_new.py', {
            'output': filepath,
            'slides': 1
        }, tools_dir)
        
        # Add slide
        result = self.run_tool('ppt_add_slide.py', {
            'file': filepath,
            'layout': 'Title and Content'
        }, tools_dir)
        
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_add_slide.py\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        assert result['data']['status'] == 'success'
        assert result['data']['total_slides'] == 2
    
    def test_set_title(self, tools_dir, temp_dir):
        """Test setting slide title."""
        # Create presentation
        filepath = temp_dir / 'title_test.pptx'
        self.run_tool('ppt_create_new.py', {
            'output': filepath,
            'slides': 1
        }, tools_dir)
        
        # Set title
        result = self.run_tool('ppt_set_title.py', {
            'file': filepath,
            'slide': 0,
            'title': 'Test Title',
            'subtitle': 'Test Subtitle'
        }, tools_dir)
        
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_set_title.py\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        assert result['data']['title'] == 'Test Title'
        assert result['data']['subtitle'] == 'Test Subtitle'
    
    def test_add_text_box_percentage(self, tools_dir, temp_dir):
        """Test adding text box with percentage positioning."""
        # Create presentation
        filepath = temp_dir / 'textbox_test.pptx'
        self.run_tool('ppt_create_new.py', {
            'output': filepath,
            'slides': 1
        }, tools_dir)
        
        # Add text box
        result = self.run_tool('ppt_add_text_box.py', {
            'file': filepath,
            'slide': 0,
            'text': 'Hello World',
            'position': {"left": "20%", "top": "30%"},
            'size': {"width": "60%", "height": "10%"},
            'font-size': 24,
            'bold': True
        }, tools_dir)
        
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_add_text_box.py\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        assert result['data']['status'] == 'success'
        assert 'Hello World' in result['data']['text']
    
    def test_add_text_box_grid(self, tools_dir, temp_dir):
        """Test adding text box with grid positioning."""
        filepath = temp_dir / 'grid_test.pptx'
        self.run_tool('ppt_create_new.py', {
            'output': filepath,
            'slides': 1
        }, tools_dir)
        
        result = self.run_tool('ppt_add_text_box.py', {
            'file': filepath,
            'slide': 0,
            'text': 'Grid Position',
            'position': {"grid": "C4"},
            'size': {"width": "25%", "height": "8%"}
        }, tools_dir)
        
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_add_text_box.py\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
    
    def test_insert_image(self, tools_dir, temp_dir):
        """Test inserting image."""
        # Create a simple test image
        try:
            from PIL import Image
            
            # Create test image
            img = Image.new('RGB', (100, 100), color='red')
            image_path = temp_dir / 'test_image.png'
            img.save(image_path)
            
            # Create presentation
            filepath = temp_dir / 'image_test.pptx'
            self.run_tool('ppt_create_new.py', {
                'output': filepath,
                'slides': 1
            }, tools_dir)
            
            # Insert image
            result = self.run_tool('ppt_insert_image.py', {
                'file': filepath,
                'slide': 0,
                'image': image_path,
                'position': {"left": "10%", "top": "10%"},
                'size': {"width": "30%", "height": "auto"},
                'alt-text': 'Test Image'
            }, tools_dir)
            
            assert result['returncode'] == 0, (
                f"\n{'='*60}\n"
                f"Tool execution failed!\n"
                f"{'='*60}\n"
                f"Tool: ppt_insert_image.py\n"
                f"Return Code: {result['returncode']}\n"
                f"\n--- STDERR ---\n{result['stderr']}\n"
                f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
                f"{'='*60}"
            )
            assert result['data']['status'] == 'success'
            assert result['data']['alt_text'] == 'Test Image'
            
        except ImportError:
            pytest.skip("Pillow not installed, skipping image test")
    
    def test_workflow_create_full_presentation(self, tools_dir, temp_dir):
        """Test complete workflow: create, add slides, set titles, add content."""
        filepath = temp_dir / 'complete_presentation.pptx'
        
        # Step 1: Create presentation
        result = self.run_tool('ppt_create_new.py', {
            'output': filepath,
            'slides': 1,
            'layout': 'Title Slide'
        }, tools_dir)
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_create_new.py (Step 1)\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        
        # Step 2: Set title on first slide
        result = self.run_tool('ppt_set_title.py', {
            'file': filepath,
            'slide': 0,
            'title': 'My Presentation',
            'subtitle': 'Created with PowerPoint Agent'
        }, tools_dir)
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_set_title.py (Step 2)\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        
        # Step 3: Add content slide
        result = self.run_tool('ppt_add_slide.py', {
            'file': filepath,
            'layout': 'Title and Content',
            'title': 'Agenda'
        }, tools_dir)
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_add_slide.py (Step 3)\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        
        # Step 4: Add text box with bullet points
        result = self.run_tool('ppt_add_text_box.py', {
            'file': filepath,
            'slide': 1,
            'text': 'Introduction\nMain Content\nConclusion',
            'position': {"left": "10%", "top": "25%"},
            'size': {"width": "80%", "height": "50%"},
            'font-size': 20
        }, tools_dir)
        assert result['returncode'] == 0, (
            f"\n{'='*60}\n"
            f"Tool execution failed!\n"
            f"{'='*60}\n"
            f"Tool: ppt_add_text_box.py (Step 4)\n"
            f"Return Code: {result['returncode']}\n"
            f"\n--- STDERR ---\n{result['stderr']}\n"
            f"\n--- STDOUT ---\n{result.get('stdout', result.get('data', ''))}\n"
            f"{'='*60}"
        )
        
        # Verify final file exists and has content
        assert filepath.exists()
        assert filepath.stat().st_size > 10000  # Should be >10KB


if __name__ == '__main__':
    # Run tests directly
    pytest.main([__file__, '-v'])
