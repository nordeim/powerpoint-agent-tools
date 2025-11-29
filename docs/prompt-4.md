Awesome job on your meticulous review, analysis and planning! Please keep up the good work in your rigorous and meticulous approach to planning and execution! Now, please meticulously review, analyze and validate the tools below.

# ORIGINAL SECTION EXTRACTED FROM `core/powerpoint_agent_core.py` (for you to take into consideration when reviewing and validating the provided tool files)  
```python
    # ========================================================================
    # VALIDATION OPERATIONS
    # ========================================================================
    
    def validate_presentation(self) -> Dict[str, Any]:
        """
        Comprehensive presentation validation.
        
        Returns:
            Validation report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        issues = {
            "empty_slides": [],
            "slides_without_titles": [],
            "fonts_used": set(),
            "large_shapes": []
        }
        
        for idx, slide in enumerate(self.prs.slides):
            # Check for empty slides
            if len(slide.shapes) == 0:
                issues["empty_slides"].append(idx)
            
            # Check for title
            has_title = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    ph_type = _get_placeholder_type_int_helper(shape.placeholder_format.type)
                    if ph_type in TITLE_PLACEHOLDER_TYPES:
                        if shape.has_text_frame and shape.text_frame.text.strip():
                            has_title = True
                            break
                
                # Collect fonts
                if hasattr(shape, 'text_frame') and shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        if para.font.name:
                            issues["fonts_used"].add(para.font.name)
            
            if not has_title:
                issues["slides_without_titles"].append(idx)
        
        issues["fonts_used"] = list(issues["fonts_used"])
        
        total_issues = (
            len(issues["empty_slides"]) +
            len(issues["slides_without_titles"])
        )
        
        return {
            "status": "issues_found" if total_issues > 0 else "valid",
            "total_issues": total_issues,
            "slide_count": len(self.prs.slides),
            "issues": issues
        }
    
    def check_accessibility(self) -> Dict[str, Any]:
        """
        Run accessibility checker.
        
        Returns:
            Accessibility report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        return AccessibilityChecker.check_presentation(self.prs)
    
    def validate_assets(self) -> Dict[str, Any]:
        """
        Run asset validator.
        
        Returns:
            Asset validation report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        return AssetValidator.validate_presentation_assets(self.prs, self.filepath)
```


---

Awesome job on your meticulous review, analysis and planning! Please keep up the good work in your rigorous and meticulous approach to planning and execution! Now, please think deeper and think harder to deeply and thoroughly explore systematically and meticulously how best to re-imagine a complete validated drop-in replacement file for each of the files, that incorporates your suggested fixes and improvements. Key consideration is to maintain compatibility with the project documents while fixing identified issues. First, review and validate the proposed updates (enhancements) against the original version again to ensure no regression after version upgrade, then meticulously create an comprehensive implementation plan to create a complete working replacement file for each of the files, include in your plan a checklist for each tool file that needs updating. Review and validate your implementation plan before proceeding meticulously to execute it. Remember to review and validate the newly generated replacement file to ensure that it is correct and accurate and that it does not contain any placeholder comment. Regression is a real threat whenever updates are made, so you need to be extremely thorough and meticulous in creating the complete updated replacement file for each of the tool files.

- Think deeply and thoroughly about how to re-imagine these tools
- Ensure compatibility with project documents
- Review and validate proposed updates against original versions to prevent regression
- Create a comprehensive implementation plan with checklists
- Review and validate the plan before executing
- Execute meticulously to create complete working replacement files
- Review and validate each replacement file to ensure correctness
- No placeholder comments - complete implementations

