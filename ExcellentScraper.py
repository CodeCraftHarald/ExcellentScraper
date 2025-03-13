#!/usr/bin/env python3
# ExcellentScraper - A web scraping tool for extracting article content to Excel

import os
import re
import time
import threading
import queue
import datetime
import random
from tkinter import filedialog
import customtkinter as ctk
from PIL import Image, ImageTk
import requests
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Set appearance mode and default color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ExcelLentScraper(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Initialize window
        self.title("ExcellentScraper")
        self.geometry("1000x800")  # Wider and taller initial size
        self.minsize(800, 600)     # Set minimum size
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # For the header
        self.grid_rowconfigure(1, weight=1)  # For the main frame
        self.grid_rowconfigure(2, weight=0)  # For the status bar
        
        # Global variables
        self.url_entries = []
        self.max_urls = 10
        self.scraped_data = []
        self.scraping_in_progress = False
        self.status_queue = queue.Queue()
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scraped_data")
        
        # Ensure the output directory exists
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        
        # Create the UI components
        self._create_ui()
        
        # Start the status update thread
        self._start_status_update_thread()
        
        # Bind keyboard shortcuts
        self.bind("<Control-r>", lambda event: self._reset_url_fields())
    
    def _create_ui(self):
        """Create the main UI components"""
        
        # Header frame with title and theme toggle
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 0))
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)
        
        # App title
        title_label = ctk.CTkLabel(
            self.header_frame,
            text="ExcellentScraper",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, sticky="w")
        
        # Theme toggle
        self.appearance_mode_menu = ctk.CTkOptionMenu(
            self.header_frame,
            values=["Dark", "Light"],
            command=self._change_appearance_mode
        )
        self.appearance_mode_menu.grid(row=0, column=1, padx=20, pady=20)
        self.appearance_mode_menu.set("Dark")
        
        # Main frame for the app content
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)  # Let content_frame expand
        
        # Create inner content frame to organize components
        content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        content_frame.grid(row=0, column=0, sticky="nsew")
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_rowconfigure(0, weight=0)  # URL frame
        content_frame.grid_rowconfigure(1, weight=0)  # Control frame
        content_frame.grid_rowconfigure(2, weight=1)  # Log frame (should expand)
        
        # URL input section - make it more compact
        self.url_frame = ctk.CTkFrame(content_frame)
        self.url_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.url_frame.grid_columnconfigure(0, weight=1)
        
        # Configure rows for URL entries (up to max_urls + additional rows for label and buttons)
        for i in range(self.max_urls + 3):  # +3 for label, spacing, and button row
            self.url_frame.grid_rowconfigure(i, weight=0)
        
        url_label = ctk.CTkLabel(
            self.url_frame,
            text="Enter URLs to scrape (all 10 fields ready):",
            font=ctk.CTkFont(size=16)
        )
        url_label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        
        # Create a container frame for the URL entries
        self.url_entries_container = ctk.CTkFrame(self.url_frame, fg_color="transparent")
        self.url_entries_container.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.url_entries_container.grid_columnconfigure(0, weight=1)  # First column
        self.url_entries_container.grid_columnconfigure(1, weight=1)  # Second column
        
        # Configure rows for the two-column layout
        for i in range((self.max_urls // 2) + (self.max_urls % 2)):  # Rows needed for an even distribution
            self.url_entries_container.grid_rowconfigure(i, weight=0)
        
        # Button frame for URL entries - position it after the URL entries
        self.url_button_frame = ctk.CTkFrame(self.url_frame, fg_color="transparent")
        self.url_button_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        
        # Add URL button
        self.add_url_button = ctk.CTkButton(
            self.url_button_frame,
            text="Add URL",
            command=lambda: self._add_url_entry(animate=True)
        )
        self.add_url_button.grid(row=0, column=0, padx=5, pady=5)
        
        # Remove URL button
        self.remove_url_button = ctk.CTkButton(
            self.url_button_frame,
            text="Remove URL",
            command=self._remove_url_entry,
            fg_color="#D35B58",
            hover_color="#C77C78"
        )
        self.remove_url_button.grid(row=0, column=1, padx=5, pady=5)
        
        # Reset button to clear all URL fields
        self.reset_button = ctk.CTkButton(
            self.url_button_frame,
            text="Reset Fields (Ctrl+R)",
            command=self._reset_url_fields,
            fg_color="#3A7EBF",
            hover_color="#5B95D0",
            width=140  # Make the button wider to fit the text
        )
        self.reset_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Create all ten URL entries at startup instead of just one
        for _ in range(self.max_urls):
            self._add_url_entry(animate=False)
        
        # Control frame
        self.control_frame = ctk.CTkFrame(content_frame)
        self.control_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        # Scrape button
        self.scrape_button = ctk.CTkButton(
            self.control_frame,
            text="Start Scraping",
            command=self._start_scraping,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.scrape_button.grid(row=0, column=0, padx=10, pady=10)
        
        # Merge button
        self.merge_button = ctk.CTkButton(
            self.control_frame,
            text="Merge Excel Files",
            command=self._merge_excel_files,
            height=40,
            font=ctk.CTkFont(size=16)
        )
        self.merge_button.grid(row=0, column=1, padx=10, pady=10)
        
        # Status and log section
        self.log_frame = ctk.CTkFrame(content_frame)
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        log_label = ctk.CTkLabel(
            self.log_frame,
            text="Status Log:",
            font=ctk.CTkFont(size=16)
        )
        log_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        # Status log text area
        self.log_text = ctk.CTkTextbox(self.log_frame, height=200)
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.configure(state="disabled")
        
        # Status bar at the bottom
        self.status_bar = ctk.CTkLabel(
            self,
            text="Ready",
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        self.status_bar.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.progress_bar.set(0)
    
    def _add_url_entry(self, animate=True):
        """Add a new URL entry field"""
        if len(self.url_entries) >= self.max_urls:
            self._update_status("Maximum number of URLs reached (10)")
            return
        
        # Create a frame for the URL entry
        entry_idx = len(self.url_entries)  # 0-based index
        
        # Calculate row and column position in the grid
        row = entry_idx // 2  # Integer division for row number
        col = entry_idx % 2   # Remainder for column (0 or 1)
        
        entry_frame = ctk.CTkFrame(self.url_entries_container, fg_color="transparent")
        entry_frame.grid(row=row, column=col, sticky="ew", padx=10, pady=5)
        entry_frame.grid_columnconfigure(1, weight=1)
        
        # Label with the entry number
        entry_label = ctk.CTkLabel(entry_frame, text=f"{entry_idx + 1}.", width=20)
        entry_label.grid(row=0, column=0, padx=(0, 5))
        
        # URL entry field
        url_entry = ctk.CTkEntry(
            entry_frame,
            placeholder_text=f"Enter URL {entry_idx + 1}"
        )
        url_entry.grid(row=0, column=1, sticky="ew", padx=5)
        
        self.url_entries.append((entry_frame, url_entry))
        
        # Update add/remove button states
        self._update_url_buttons()
        
        # Add a little sparkle animation only if requested
        if animate and entry_idx > 0:  # Don't animate the first entry
            self._animate_entry_addition(entry_frame)
    
    def _animate_entry_addition(self, frame):
        """Create a small animation when adding a new entry"""
        original_color = frame.cget("fg_color")
        highlight_color = "#2a6496" if ctk.get_appearance_mode() == "Dark" else "#a2d2ff"
        
        def _animate_step(step=0, max_steps=10):
            if step <= max_steps:
                # Gradually fade from highlight color to original
                blend_factor = step / max_steps
                r1, g1, b1 = [int(highlight_color[1:3], 16), int(highlight_color[3:5], 16), int(highlight_color[5:7], 16)]
                
                # Handle the case where original_color might be "transparent"
                if original_color == "transparent":
                    r2, g2, b2 = [40, 40, 40] if ctk.get_appearance_mode() == "Dark" else [240, 240, 240]
                else:
                    r2, g2, b2 = [int(original_color[1:3], 16), int(original_color[3:5], 16), int(original_color[5:7], 16)]
                
                r = int(r1 * (1 - blend_factor) + r2 * blend_factor)
                g = int(g1 * (1 - blend_factor) + g2 * blend_factor)
                b = int(b1 * (1 - blend_factor) + b2 * blend_factor)
                
                current_color = f"#{r:02x}{g:02x}{b:02x}"
                frame.configure(fg_color=current_color)
                self.after(30, lambda: _animate_step(step + 1, max_steps))
            else:
                frame.configure(fg_color=original_color)
        
        _animate_step()
    
    def _remove_url_entry(self):
        """Remove the last URL entry field"""
        if not self.url_entries:
            return
        
        # Get the last entry and remove it
        entry_frame, _ = self.url_entries.pop()
        entry_frame.destroy()
        
        # Update add/remove button states
        self._update_url_buttons()
    
    def _reset_url_fields(self):
        """Clear all URL entry fields without removing them"""
        # Check if there are entries to clear
        if not self.url_entries:
            return
            
        # Clear each URL entry
        for _, url_entry in self.url_entries:
            url_entry.delete(0, 'end')  # Clear the entry
        
        # Show a brief status message
        self._update_status("All URL fields have been cleared")
        
        # Add a visual feedback effect
        self._flash_url_container()
    
    def _flash_url_container(self):
        """Provide visual feedback that the URL fields have been reset"""
        original_color = self.url_entries_container.cget("fg_color")
        highlight_color = "#2a6496" if ctk.get_appearance_mode() == "Dark" else "#a2d2ff"
        
        def _flash_step(step=0, max_steps=10):
            if step <= max_steps:
                # Gradually fade from highlight color to original
                blend_factor = step / max_steps
                r1, g1, b1 = [int(highlight_color[1:3], 16), int(highlight_color[3:5], 16), int(highlight_color[5:7], 16)]
                
                # Handle the case where original_color might be "transparent"
                if original_color == "transparent":
                    r2, g2, b2 = [40, 40, 40] if ctk.get_appearance_mode() == "Dark" else [240, 240, 240]
                else:
                    r2, g2, b2 = [int(original_color[1:3], 16), int(original_color[3:5], 16), int(original_color[5:7], 16)]
                
                r = int(r1 * (1 - blend_factor) + r2 * blend_factor)
                g = int(g1 * (1 - blend_factor) + g2 * blend_factor)
                b = int(b1 * (1 - blend_factor) + b2 * blend_factor)
                
                current_color = f"#{r:02x}{g:02x}{b:02x}"
                self.url_entries_container.configure(fg_color=current_color)
                self.after(30, lambda: _flash_step(step + 1, max_steps))
            else:
                self.url_entries_container.configure(fg_color=original_color)
        
        _flash_step()
    
    def _update_url_buttons(self):
        """Update the state of the add/remove URL buttons"""
        if len(self.url_entries) >= self.max_urls:
            self.add_url_button.configure(state="disabled")
        else:
            self.add_url_button.configure(state="normal")
        
        if not self.url_entries:
            self.remove_url_button.configure(state="disabled")
        else:
            self.remove_url_button.configure(state="normal")
    
    def _change_appearance_mode(self, mode):
        """Change the app's appearance mode (dark/light)"""
        ctk.set_appearance_mode(mode.lower())
        self._update_status(f"Theme changed to {mode} mode")
    
    def _update_status(self, message):
        """Update the status bar and log with a message"""
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{current_time}] {message}"
        
        # Add to the queue (thread-safe)
        self.status_queue.put(formatted_message)
    
    def _start_status_update_thread(self):
        """Start a thread to update the status display"""
        def update_status_display():
            while True:
                try:
                    # Check if there are any status updates in the queue
                    if not self.status_queue.empty():
                        message = self.status_queue.get()
                        
                        # Update the status bar
                        self.status_bar.configure(text=message.split("] ")[1])
                        
                        # Update the log text
                        self.log_text.configure(state="normal")
                        self.log_text.insert("end", message + "\n")
                        self.log_text.see("end")
                        self.log_text.configure(state="disabled")
                    
                    # Sleep briefly to reduce CPU usage
                    time.sleep(0.1)
                except Exception as e:
                    print(f"Error in status update thread: {e}")
        
        # Start the thread
        threading.Thread(target=update_status_display, daemon=True).start()
    
    def _collect_urls(self):
        """Collect URLs from the entry fields"""
        urls = []
        for _, url_entry in self.url_entries:
            url = url_entry.get().strip()
            if url:
                # Add http:// if it's missing
                if not url.startswith(("http://", "https://")):
                    url = "https://" + url
                urls.append(url)
        return urls
    
    def _start_scraping(self):
        """Start the scraping process"""
        if self.scraping_in_progress:
            self._update_status("Scraping already in progress")
            return
        
        urls = self._collect_urls()
        if not urls:
            self._update_status("No URLs entered. Please enter at least one URL")
            return
        
        # Disable buttons during scraping
        self.scrape_button.configure(state="disabled", text="Scraping...")
        self.merge_button.configure(state="disabled")
        self.scraping_in_progress = True
        
        # Reset progress bar
        self.progress_bar.set(0)
        
        # Start the scraping thread
        threading.Thread(target=self._scrape_urls, args=(urls,), daemon=True).start()
    
    def _scrape_urls(self, urls):
        """Scrape the provided URLs in a thread"""
        self._update_status(f"Starting to scrape {len(urls)} URLs...")
        self.scraped_data = []
        
        # Initialize webdriver lazily only when needed
        driver = None
        
        for i, url in enumerate(urls):
            try:
                # Update progress
                progress = (i) / len(urls)
                self.progress_bar.set(progress)
                
                self._update_status(f"Scraping URL {i+1}/{len(urls)}: {url}")
                
                # First try with requests and BeautifulSoup
                try:
                    article_data = self._scrape_with_beautifulsoup(url)
                    self._update_status(f"Successfully scraped with BeautifulSoup: {url}")
                except Exception as bs_error:
                    self._update_status(f"BeautifulSoup failed for {url}, trying Selenium: {str(bs_error)}")
                    
                    # Initialize Selenium if not already done
                    if driver is None:
                        self._update_status("Initializing Selenium WebDriver...")
                        chrome_options = Options()
                        chrome_options.add_argument("--headless")
                        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
                    
                    # Try with Selenium
                    article_data = self._scrape_with_selenium(driver, url)
                    self._update_status(f"Successfully scraped with Selenium: {url}")
                
                # Add the scraped data
                self.scraped_data.append(article_data)
                
                # Add some randomized delay to prevent rate limiting
                time.sleep(random.uniform(0.5, 2.0))
                
            except Exception as e:
                self._update_status(f"Error scraping {url}: {str(e)}")
        
        # Close the driver if it was initialized
        if driver:
            driver.quit()
            self._update_status("Closed Selenium WebDriver")
        
        # Export the data to Excel
        if self.scraped_data:
            filename = self._export_to_excel()
            self._update_status(f"Exported data to: {filename}")
        else:
            self._update_status("No data was scraped")
        
        # Complete
        self.progress_bar.set(1.0)
        self._update_status("Scraping completed - You can now reset the fields for a new batch")
        
        # Re-enable buttons
        self.scrape_button.configure(state="normal", text="Start Scraping")
        self.merge_button.configure(state="normal")
        self.scraping_in_progress = False
    
    def _scrape_with_beautifulsoup(self, url):
        """Scrape a URL using requests and BeautifulSoup"""
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Pragma': 'no-cache',
            'Cache-Control': 'no-cache',
        }
        
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        # Try to detect encoding, defaulting to UTF-8
        if response.encoding is None or response.encoding == 'ISO-8859-1':
            # Requests sometimes incorrectly detects ISO-8859-1
            possible_encoding = response.apparent_encoding
            if possible_encoding and possible_encoding.lower() != 'iso-8859-1':
                response.encoding = possible_encoding
        
        # Use html.parser first, but fall back to lxml if available, and html5lib as a last resort
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Check if the page content was properly parsed
        if not soup.find('body') or len(soup.get_text(strip=True)) < 100:
            try:
                soup = BeautifulSoup(response.text, 'lxml')
            except:
                try:
                    soup = BeautifulSoup(response.text, 'html5lib')
                except:
                    pass  # Stick with html.parser
        
        # Extract the title
        title = self._extract_title(soup)
        
        # Extract headings (improved filtering)
        headings = []
        for heading in soup.find_all(['h1', 'h2', 'h3']):
            text = heading.get_text(strip=True)
            if text and len(text) > 3:  # Filter out very short or empty headings
                # Check if heading isn't just navigation or generic text
                if not any(nav_word in text.lower() for nav_word in ['menu', 'navigation', 'search', 'login', 'sign in']):
                    headings.append(text)
        
        # If no headings were found, use the title as the first heading
        if not headings and title:
            headings = [title]
        
        # Extract the main content, trying different approaches
        content = self._extract_article_content(soup)
        
        # Return the data
        return {
            'url': url,
            'title': title,
            'headings': headings,
            'content': content,
            'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    
    def _scrape_with_selenium(self, driver, url):
        """Scrape a URL using Selenium"""
        driver.get(url)
        
        # Wait for the page to load (increased timeout and better detection)
        try:
            # First wait for body to exist
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # Then wait for content to load - try common article content selectors
            content_selectors = [
                "article", 
                ".article", 
                ".post", 
                ".content",
                ".entry-content", 
                ".article-content",
                "#content", 
                ".main-content"
            ]
            
            # Try each selector for a short time
            for selector in content_selectors:
                try:
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    self._update_status(f"Found content using selector: {selector}")
                    break
                except:
                    continue
                    
            # As a fallback, just wait a bit for any dynamic content to load
            # This helps with JavaScript-heavy sites
            time.sleep(2)
            
        except Exception as e:
            self._update_status(f"Warning: Timeout waiting for page to fully load: {str(e)}")
        
        # Extract the title
        title = driver.title
        
        # Get the page source and parse it with BeautifulSoup
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # Extract headings
        headings = []
        for heading in soup.find_all(['h1', 'h2', 'h3']):
            text = heading.get_text(strip=True)
            if text and len(text) > 3:  # Filter out very short or empty headings
                # Check if heading isn't just navigation or generic text
                if not any(nav_word in text.lower() for nav_word in ['menu', 'navigation', 'search', 'login', 'sign in']):
                    headings.append(text)
        
        # If no headings were found, use the title as the first heading
        if not headings and title:
            headings = [title]
        
        # Extract the main content
        content = self._extract_article_content(soup)
        
        # Return the data
        return {
            'url': url,
            'title': title,
            'headings': headings,
            'content': content,
            'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    
    def _extract_title(self, soup):
        """Extract the title of the article"""
        title_candidates = []
        
        # Try different title sources in order of reliability for article content
        
        # 1. Schema.org article headline
        article_headline = soup.find('meta', {'itemprop': 'headline'})
        if article_headline and article_headline.get('content'):
            title_candidates.append(article_headline['content'].strip())
        
        # 2. Open Graph title
        og_title = soup.find('meta', property='og:title') or soup.find('meta', attrs={'name': 'og:title'})
        if og_title and og_title.get('content'):
            title_candidates.append(og_title['content'].strip())
            
        # 3. Twitter card title
        twitter_title = soup.find('meta', attrs={'name': 'twitter:title'})
        if twitter_title and twitter_title.get('content'):
            title_candidates.append(twitter_title['content'].strip())
        
        # 4. Main heading
        main_heading = soup.find('h1')
        if main_heading and main_heading.text.strip():
            # Make sure it's not a site name or navigation
            heading_text = main_heading.text.strip()
            if len(heading_text.split()) > 1 and not any(nav_term in heading_text.lower() for nav_term in ['home', 'menu', 'navigation']):
                title_candidates.append(heading_text)
        
        # 5. Page title tag
        title_tag = soup.find('title')
        if title_tag and title_tag.string:
            page_title = title_tag.string.strip()
            
            # Try to remove site name from title
            if ' | ' in page_title:
                page_title = page_title.split(' | ')[0].strip()
            elif ' - ' in page_title:
                page_title = page_title.split(' - ')[0].strip()
            elif ' – ' in page_title:
                page_title = page_title.split(' – ')[0].strip()
                
            title_candidates.append(page_title)
            
        # 6. Any other h1 if we still don't have candidates
        if not title_candidates:
            for h1 in soup.find_all('h1'):
                if h1.text.strip() and len(h1.text.strip()) > 10:  # Require a minimum length
                    title_candidates.append(h1.text.strip())
                    break
        
        # Choose the best title from candidates
        if title_candidates:
            # Prefer longer titles as they are typically more descriptive
            # But not too long (avoid full paragraphs)
            filtered_candidates = [t for t in title_candidates if 3 < len(t.split()) < 20]
            
            if filtered_candidates:
                return max(filtered_candidates, key=len)
            else:
                # If no good candidates after filtering, take the first one
                return title_candidates[0]
        
        return "No title found"
    
    def _extract_article_content(self, soup):
        """Extract the main content of the article"""
        # Try to identify the main article content by common patterns
        article_candidates = []
        
        # Look for article tag
        article = soup.find('article')
        if article:
            article_candidates.append(article)
        
        # Look for common content div IDs and classes
        content_selectors = [
            '#content', '.content',
            '#main', '.main',
            '#article', '.article',
            '#post', '.post',
            '.post-content', '.entry-content',
            '.article-body', '.story-body',
            '.article-content', '.entry',
            '.main-content', '.page-content',
            '.story', '.blog-post',
            '.cms-content', '.node-content',
            '.rich-text', '.article__body',
            '.entry__content', '.post__content'
        ]
        
        for selector in content_selectors:
            elements = soup.select(selector)
            if elements:
                article_candidates.extend(elements)
        
        # Find the candidate with the most text content, excluding navigation, ads, etc.
        if article_candidates:
            # Clean up candidates before measuring text length
            for candidate in article_candidates:
                # Create a deep copy to work with
                candidate_copy = candidate
                
                # Remove unwanted elements from the copy
                for unwanted in candidate_copy.find_all(['script', 'style', 'iframe', 'nav', 'footer', 'header', 
                                                       'aside', '.sidebar', '.widget', '.ad', '.advertisement',
                                                       '.social', '.comments', '.related', '.recommended',
                                                       '.newsletter', '.promo']):
                    unwanted.decompose()
            
            # Sort by text length and pick the longest
            article_candidates.sort(key=lambda x: len(x.get_text(strip=True)), reverse=True)
            main_content = article_candidates[0]
            
            # Clean up the content more thoroughly
            for tag in main_content.find_all(['script', 'style', 'iframe', 'nav', 'footer', 'header', 
                                           'button', '.nav', '.menu', '.sidebar', '.widget', '.ad', 
                                           '.social-share', '.share-buttons', '.comments', '.comment-section',
                                           '.related-posts', '.recommended-articles', '.newsletter-signup']):
                tag.decompose()
            
            # Get all paragraphs from main content
            paragraphs = []
            for p in main_content.find_all('p'):
                text = p.get_text(strip=True)
                # Filter out short or likely non-article paragraphs
                if text and len(text.split()) > 4 and not re.match(r'^(share|posted by|written by|author:|date:|published:).*$', text.lower()):
                    paragraphs.append(text)
            
            if paragraphs:
                return "\n\n".join(paragraphs)
            
            # Fallback to full text if paragraph extraction failed
            content = main_content.get_text(separator="\n").strip()
            content = re.sub(r'\n{3,}', '\n\n', content)  # Remove excessive newlines
            content = re.sub(r'[\t ]+', ' ', content)     # Normalize whitespace
            return content
        
        # Fallback: extract all paragraph text
        paragraphs = []
        for p in soup.find_all('p'):
            text = p.get_text(strip=True)
            # More aggressive filtering for potential non-content paragraphs
            if text and len(text.split()) > 5 and not any(phrase in text.lower() for phrase in 
                                                       ['cookie', 'privacy policy', 'terms of service', 
                                                        'copyright', 'all rights reserved', 'newsletter', 
                                                        'sign up', 'subscribe']):
                paragraphs.append(text)
        
        if paragraphs:
            return "\n\n".join(paragraphs)
        
        # Last resort: get the main text content while filtering out common non-content areas
        body = soup.find('body')
        if body:
            # Remove non-content elements
            non_content_selectors = [
                'header', 'footer', 'nav', 'aside', 
                '.sidebar', '.widget', '.comments', '.ad', 
                '.advertisement', '.menu', '.navigation', 
                '.social', '.share', '.related', '.recommended'
            ]
            
            for selector in non_content_selectors:
                for element in body.select(selector):
                    element.decompose()
            
            # Also remove script, style, etc.
            for tag in body.find_all(['script', 'style', 'iframe', 'noscript']):
                tag.decompose()
            
            # Extract and clean the text
            content = body.get_text(separator="\n").strip()
            content = re.sub(r'\n{3,}', '\n\n', content)  # Remove excessive newlines
            content = re.sub(r'[\t ]+', ' ', content)     # Normalize whitespace
            
            # Try to find the part of the content with the highest content density
            lines = content.split('\n')
            if len(lines) > 20:  # If content is long enough to be worth analyzing
                # Find longest consecutive group of non-empty lines (likely the article)
                best_start = 0
                best_length = 0
                current_start = 0
                current_length = 0
                
                for i, line in enumerate(lines):
                    if line.strip():
                        if current_length == 0:
                            current_start = i
                        current_length += 1
                    else:
                        if current_length > best_length:
                            best_start = current_start
                            best_length = current_length
                        current_length = 0
                
                # Handle the case where the best segment is at the end
                if current_length > best_length:
                    best_start = current_start
                    best_length = current_length
                
                # Extract the best content segment if it's significant
                if best_length > 5:
                    content = '\n'.join(lines[best_start:best_start + best_length])
            
            return content
        
        return "No content found"
    
    def _export_to_excel(self):
        """Export the scraped data to an Excel file"""
        # Create a timestamp for the filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.output_dir, f"scraped_data_{timestamp}.xlsx")
        
        # Prepare the data for Excel
        excel_data = []
        
        for article in self.scraped_data:
            # Basic columns
            row = [
                article['timestamp'],
                article['url'],
                article['headings'][0] if article['headings'] else "No heading",
                article['content']
            ]
            
            # Add additional headings as separate columns
            for i, heading in enumerate(article['headings'][1:], 1):
                # Extend the row if needed
                while len(row) <= 3 + i:
                    row.append("")
                row[3 + i] = heading
            
            excel_data.append(row)
        
        # Define column names
        columns = ['Timestamp', 'URL', 'First Heading', 'Content']
        
        # Find the maximum number of headings
        max_headings = 1  # Always have at least the first heading
        for article in self.scraped_data:
            max_headings = max(max_headings, len(article['headings']))
        
        # Add additional heading columns
        for i in range(1, max_headings):
            columns.append(f'Heading {i+1}')
        
        # Create a DataFrame
        df = pd.DataFrame(excel_data, columns=columns)
        
        # Write to Excel
        df.to_excel(filename, index=False, engine='openpyxl')
        
        return filename
    
    def _merge_excel_files(self):
        """Merge multiple Excel files"""
        # Open file dialog to select the files
        files = filedialog.askopenfilenames(
            title="Select Excel files to merge",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self.output_dir
        )
        
        if not files:
            self._update_status("No files selected for merging")
            return
        
        self._update_status(f"Selected {len(files)} files for merging")
        
        # Ask for the output file
        output_file = filedialog.asksaveasfilename(
            title="Save merged Excel file as",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self.output_dir,
            defaultextension=".xlsx"
        )
        
        if not output_file:
            self._update_status("Merge operation cancelled")
            return
        
        try:
            # Start with an empty DataFrame or the first file if it exists
            if os.path.exists(output_file):
                merged_df = pd.read_excel(output_file, engine='openpyxl')
                self._update_status(f"Loaded existing file: {output_file}")
            else:
                merged_df = pd.DataFrame()
            
            # Merge each file
            for file in files:
                try:
                    df = pd.read_excel(file, engine='openpyxl')
                    self._update_status(f"Reading file: {os.path.basename(file)}")
                    
                    # Append the data
                    merged_df = pd.concat([merged_df, df], ignore_index=True)
                    
                    # Add a little delay for visual effect
                    time.sleep(0.2)
                    
                except Exception as e:
                    self._update_status(f"Error reading {os.path.basename(file)}: {str(e)}")
            
            # Save the merged DataFrame
            merged_df.to_excel(output_file, index=False, engine='openpyxl')
            self._update_status(f"Successfully merged files into: {output_file}")
            
            # Show a success animation
            self._animate_merge_success()
            
        except Exception as e:
            self._update_status(f"Error merging files: {str(e)}")
    
    def _animate_merge_success(self):
        """Animate a success message after merging files"""
        success_frame = ctk.CTkFrame(self, corner_radius=10)
        success_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        success_label = ctk.CTkLabel(
            success_frame,
            text="✨ Files Merged Successfully! ✨",
            font=ctk.CTkFont(size=18, weight="bold"),
            padx=20,
            pady=20
        )
        success_label.grid(row=0, column=0)
        
        # Fade out animation
        def fade_out(alpha=1.0):
            if alpha > 0:
                success_frame.attributes("-alpha", alpha)
                self.after(50, lambda: fade_out(alpha - 0.05))
            else:
                success_frame.destroy()
        
        # Start fade out after 2 seconds
        self.after(2000, fade_out)


if __name__ == "__main__":
    app = ExcelLentScraper()
    app.mainloop() 