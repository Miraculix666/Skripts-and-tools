import re
import pytest
from playwright.sync_api import Page, expect

def get_html_parts():
    with open('maybe_unneeded_files.ps1', 'r', encoding='utf-8') as f:
        content = f.read()

    style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
    script_match = re.search(r'<script>(.*?)</script>', content, re.DOTALL)

    style = style_match.group(1) if style_match else ""
    script = script_match.group(1) if script_match else ""

    return style, script

def test_collapse_all(page: Page):
    """
    Test that collapseAll() properly hides archive items
    and ensures non-archive items are visible.
    """
    style, script = get_html_parts()

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>{style}</style>
    </head>
    <body>
        <!-- Setup various states -->
        <div class="file-item" id="normal-visible" style="display: flex;">Normal File 1</div>
        <div class="file-item" id="normal-hidden" style="display: none;">Normal File 2</div>
        <div class="file-item archive" id="archive-hidden" style="display: none;">Archive File 1</div>
        <div class="file-item archive" id="archive-visible" style="display: flex;">Archive File 2</div>

        <script>{script}</script>
    </body>
    </html>
    """

    page.set_content(html_content)

    # Pre-assertions to ensure test data is setup correctly
    assert page.locator('#normal-visible').evaluate("el => el.style.display") == 'flex'
    assert page.locator('#normal-hidden').evaluate("el => el.style.display") == 'none'
    assert page.locator('#archive-hidden').evaluate("el => el.style.display") == 'none'
    assert page.locator('#archive-visible').evaluate("el => el.style.display") == 'flex'

    # Call collapseAll
    page.evaluate("collapseAll()")

    # Assertions post-collapseAll()

    # Normal files should ALWAYS be display: flex
    assert page.locator('#normal-visible').evaluate("el => el.style.display") == 'flex'
    assert page.locator('#normal-hidden').evaluate("el => el.style.display") == 'flex'

    # Archive files should ALWAYS be display: none
    assert page.locator('#archive-hidden').evaluate("el => el.style.display") == 'none'
    assert page.locator('#archive-visible').evaluate("el => el.style.display") == 'none'
