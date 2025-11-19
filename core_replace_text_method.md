```python
    def replace_text(self, find: str, replace: str, match_case: bool = False) -> int:
        """Find and replace text - ultra-simple version."""
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        count = 0
        
        for slide in self.prs.slides:
            for shape in slide.shapes:
                try:
                    # Try to get text - if it fails, skip this shape
                    old_text = shape.text
                except (AttributeError, Exception):
                    continue
                
                # Skip empty text
                if not old_text:
                    continue
                
                # Replace based on case sensitivity
                if match_case:
                    if find in old_text:
                        new_text = old_text.replace(find, replace)
                        shape.text = new_text
                        count += old_text.count(find)
                else:
                    if find.lower() in old_text.lower():
                        pattern = re.compile(re.escape(find), re.IGNORECASE)
                        matches = re.findall(pattern, old_text)
                        new_text = pattern.sub(replace, old_text)
                        shape.text = new_text
                        count += len(matches)
        
        return count
```
