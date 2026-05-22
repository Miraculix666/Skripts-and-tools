import pytest
import re
import os
from playwright.sync_api import sync_playwright

def get_script_and_style():
    # Resolve the path dynamically relative to this script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    ps1_path = os.path.join(current_dir, '..', 'maybe_unneeded_files.ps1')

    with open(ps1_path, 'r', encoding='utf-8') as f:
        content = f.read()

    style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
    script_match = re.search(r'<script>(.*?)</script>', content, re.DOTALL)

    style = style_match.group(1) if style_match else ''
    script = script_match.group(1) if script_match else ''

    return style, script

def test_expandAll():
    style, script = get_script_and_style()

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>{style}</style>
    </head>
    <body>
        <div class="file-list">
            <div class="file-item" style="display: none;">Item 1</div>
            <div class="file-item archive" style="display: none;">Item 2 (Archive)</div>
            <div class="file-item">Item 3 (Visible)</div>
        </div>
        <script>{script}</script>
    </body>
    </html>
    """

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(html_content)

        # Verify initial state
        assert page.locator('.file-item').nth(0).evaluate('el => el.style.display') == 'none'
        assert page.locator('.file-item').nth(1).evaluate('el => el.style.display') == 'none'

        # Execute expandAll
        page.evaluate('expandAll()')

        # Verify expandAll behavior
        assert page.locator('.file-item').nth(0).evaluate('el => el.style.display') == 'flex'
        assert page.locator('.file-item').nth(1).evaluate('el => el.style.display') == 'flex'
        assert page.locator('.file-item').nth(2).evaluate('el => el.style.display') == 'flex'

        browser.close()

def test_toggleArchive():
    style, script = get_script_and_style()

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>{style}</style>
    </head>
    <body>
        <div class="file-list">
            <div class="file-item" style="display: none;">Item 1</div>
            <div class="file-item archive" style="display: none;">Item 2 (Archive)</div>
            <div class="file-item archive" style="display: flex;">Item 3 (Archive)</div>
            <div class="file-item">Item 4 (Visible)</div>
        </div>
        <script>{script}</script>
    </body>
    </html>
    """

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(html_content)

        # Verify initial state
        assert page.locator('.archive').nth(0).evaluate('el => el.style.display') == 'none'
        assert page.locator('.archive').nth(1).evaluate('el => el.style.display') == 'flex'

        # Execute toggleArchive
        page.evaluate('toggleArchive()')

        # Verify toggleArchive behavior
        assert page.locator('.archive').nth(0).evaluate('el => el.style.display') == 'flex'
        assert page.locator('.archive').nth(1).evaluate('el => el.style.display') == 'none'

        browser.close()

def test_collapseAll():
    style, script = get_script_and_style()

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>{style}</style>
    </head>
    <body>
        <div class="file-list">
            <div class="file-item" style="display: none;">Item 1</div>
            <div class="file-item archive" style="display: none;">Item 2 (Archive)</div>
            <div class="file-item archive" style="display: flex;">Item 3 (Archive)</div>
            <div class="file-item" style="display: flex;">Item 4 (Visible)</div>
        </div>
        <script>{script}</script>
    </body>
    </html>
    """

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(html_content)

        # Execute collapseAll
        page.evaluate('collapseAll()')

        # Verify collapseAll behavior
        # Normal items should be flex
        assert page.locator('.file-item:not(.archive)').nth(0).evaluate('el => el.style.display') == 'flex'
        assert page.locator('.file-item:not(.archive)').nth(1).evaluate('el => el.style.display') == 'flex'

        # Archive items should be none
        assert page.locator('.archive').nth(0).evaluate('el => el.style.display') == 'none'
        assert page.locator('.archive').nth(1).evaluate('el => el.style.display') == 'none'

        browser.close()
