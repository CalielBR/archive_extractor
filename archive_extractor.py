"""
Archive Extractor with Graphical Interface.

Script to extract compressed files (ZIP, RAR, 7Z) with graphical interface,
support for split files and progress display.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import time
import zipfile
from typing import Optional, Callable, List
import rarfile
import py7zr


class ExtractionStatus:
    """Stores extraction progress and statistics."""

    def __init__(self):
        self.total_files = 0
        self.total_size_gb = 0.0
        self.current_file_index = 0
        self.current_file_progress = 0
        self.current_file_total = 0


class ArchiveExtractor:
    """Main class for handling compressed files."""

    def __init__(self):
        self.status = ExtractionStatus()

    def join_zip_parts(self, folder: str, base_name: str) -> Optional[str]:
        """
        Join parts of a split zip file into a single file.

        Args:
            folder: Directory containing .zip.xxx segments.
            base_name: Base filename (before .zip.xxx).

        Returns:
            Path to the merged zip file, or None if no parts found.
        """
        parts = []
        i = 1

        # Find all numbered parts
        while True:
            part_path = os.path.join(folder, f"{base_name}.zip.{i:03d}")
            if not os.path.exists(part_path):
                break
            parts.append(part_path)
            i += 1

        if not parts:
            return None

        # Create unified file
        merged_path = os.path.join(folder, f"{base_name}_merged.zip")
        try:
            with open(merged_path, "wb") as output_file:
                for part in parts:
                    with open(part, "rb") as input_file:
                        output_file.write(input_file.read())
            return merged_path
        except IOError as e:
            print(f"Error joining parts: {e}")
            return None

    def extract_archive(self, archive_path: str, destination_folder: str,
                        progress_callback: Optional[Callable] = None) -> str:
        """
        Extract a compressed file (zip, rar, 7z or zip parts).

        Args:
            archive_path: Input file path.
            destination_folder: Destination folder for extracted files.
            progress_callback: Function to update progress.

        Returns:
            Success or error message from extraction.
        """
        if not os.path.exists(archive_path):
            return f"❌ File not found: {archive_path}"

        file_size_bytes = os.path.getsize(archive_path)
        file_size_gb = file_size_bytes / (1024 ** 3)
        file_name = os.path.basename(archive_path)

        try:
            archive_path_lower = archive_path.lower()

            if archive_path_lower.endswith(".zip"):
                return self._extract_zip(archive_path, destination_folder,
                                         file_size_gb, progress_callback)

            elif archive_path_lower.endswith(".rar"):
                return self._extract_rar(archive_path, destination_folder,
                                         file_size_gb, progress_callback)

            elif archive_path_lower.endswith(".7z"):
                return self._extract_7z(archive_path, destination_folder,
                                        file_size_gb, progress_callback)

            elif ".zip." in archive_path_lower:
                return self._extract_split_zip(archive_path, destination_folder,
                                               file_size_gb, progress_callback)

            else:
                return (f"❌ Format not supported: {file_name} "
                        f"({file_size_gb:.2f} GB)")

        except (zipfile.BadZipFile, rarfile.Error,
                py7zr.exceptions.Bad7zFile, IOError, OSError) as e:
            return f"❌ Error extracting {file_name} ({file_size_gb:.2f} GB): {str(e)}"

    def _extract_zip(self, archive_path: str, destination_folder: str,
                     file_size_gb: float, progress_callback: Optional[Callable]) -> str:
        """Extract ZIP files."""
        with zipfile.ZipFile(archive_path, "r") as zip_ref:
            file_list = zip_ref.namelist()
            total_files = len(file_list)

            for i, file_name in enumerate(file_list, 1):
                zip_ref.extract(file_name, destination_folder)
                if progress_callback:
                    progress_callback(i, total_files, file_size_gb)

        return (f"✅ Successfully extracted: {os.path.basename(archive_path)} "
                f"({file_size_gb:.2f} GB)")

    def _extract_rar(self, archive_path: str, destination_folder: str,
                     file_size_gb: float, progress_callback: Optional[Callable]) -> str:
        """Extract RAR files."""
        with rarfile.RarFile(archive_path, "r") as rar_ref:
            file_list = rar_ref.namelist()
            total_files = len(file_list)

            for i, file_name in enumerate(file_list, 1):
                rar_ref.extract(file_name, destination_folder)
                if progress_callback:
                    progress_callback(i, total_files, file_size_gb)

        return (f"✅ Successfully extracted: {os.path.basename(archive_path)} "
                f"({file_size_gb:.2f} GB)")

    def _extract_7z(self, archive_path: str, destination_folder: str,
                    file_size_gb: float, progress_callback: Optional[Callable]) -> str:
        """Extract 7Z files."""
        with py7zr.SevenZipFile(archive_path, "r") as seven_ref:
            file_list = seven_ref.getnames()
            total_files = len(file_list)

            for i, file_name in enumerate(file_list, 1):
                seven_ref.extract(targets=[file_name], path=destination_folder)
                if progress_callback:
                    progress_callback(i, total_files, file_size_gb)

        return (f"✅ Successfully extracted: {os.path.basename(archive_path)} "
                f"({file_size_gb:.2f} GB)")

    def _extract_split_zip(self, archive_path: str, destination_folder: str,
                           file_size_gb: float, progress_callback: Optional[Callable]) -> str:
        """Extract split ZIP files."""
        folder = os.path.dirname(archive_path)
        base_name = os.path.basename(archive_path).split(".zip.")[0]

        merged_path = self.join_zip_parts(folder, base_name)
        if not merged_path:
            return (f"❌ Parts not found: {os.path.basename(archive_path)} "
                    f"({file_size_gb:.2f} GB)")

        try:
            result = self._extract_zip(merged_path, destination_folder,
                                       file_size_gb, progress_callback)
            # Remove temporary file
            if os.path.exists(merged_path):
                os.remove(merged_path)
            return result
        except (zipfile.BadZipFile, IOError, OSError) as e:
            return f"❌ Error extracting parts: {str(e)}"


class ExtractionGUI:
    """Graphical interface for file extraction."""

    def __init__(self, root):
        self.root = root
        self.extractor = ArchiveExtractor()
        self.countdown_label = None
        self.countdown_window = None
        self.countdown_running = False
        self.setup_gui()

    def setup_gui(self):
        """Configure the graphical interface."""
        self.root.title("File Extractor")
        self.root.geometry("800x600")

        # Main frame
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Select Files button (GREEN with larger font)
        select_files_btn = tk.Button(
            main_frame,
            text="📁 Select Files",
            command=self.select_files,
            width=20,
            font=("Arial", 12, "bold"),
            bg="#4CAF50",  # Green
            fg="white",
            relief=tk.RAISED,
            borderwidth=2
        )
        select_files_btn.grid(row=0, column=0, sticky="ew", pady=5)

        # File list
        self.file_listbox = tk.Listbox(main_frame, width=80, height=10)
        self.file_listbox.grid(row=1, column=0, columnspan=3,
                               pady=5, sticky="nsew")

        # Destination frame
        dest_frame = tk.Frame(main_frame)
        dest_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=5)

        # Select Destination button (PURPLE with larger font)
        select_dest_btn = tk.Button(
            dest_frame,
            text="📂 Select Destination Folder",
            command=self.select_destination,
            width=22,
            font=("Arial", 11, "bold"),
            bg="#9C27B0",  # Purple
            fg="white",
            relief=tk.RAISED,
            borderwidth=2
        )
        select_dest_btn.pack(side=tk.LEFT)

        self.destination_var = tk.StringVar()
        tk.Entry(dest_frame, textvariable=self.destination_var,
                 width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Extract button (blue for contrast)
        extract_btn = tk.Button(
            dest_frame,
            text="🚀 Extract",
            command=self.start_extraction,
            width=15,
            font=("Arial", 11, "bold"),
            bg="#2196F3",  # Blue
            fg="white",
            relief=tk.RAISED,
            borderwidth=2
        )
        extract_btn.pack(side=tk.RIGHT)

        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, length=500)
        self.progress_bar.grid(row=3, column=0, columnspan=3,
                               pady=5, sticky="ew")

        # Status label
        self.status_label = tk.Label(main_frame,
                                     text="Status: waiting for extraction…",
                                     font=("Arial", 10))
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

        # Results area
        self.results_text = scrolledtext.ScrolledText(
            main_frame, width=80, height=10)
        self.results_text.grid(row=5, column=0, columnspan=3,
                               pady=10, sticky="nsew")

        # Configure weights for expansion
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def select_files(self):
        """Open dialog to select files."""
        files = filedialog.askopenfilenames(
            title="Select files to extract",
            filetypes=[
                ("Compressed files", "*.zip *.rar *.7z *.001 *.002 *.003"),
                ("All files", "*.*")
            ]
        )
        self.file_listbox.delete(0, tk.END)
        for file in files:
            self.file_listbox.insert(tk.END, file)

    def select_destination(self):
        """Select destination folder."""
        folder = filedialog.askdirectory(title="Select destination folder")
        if folder:
            self.destination_var.set(folder)

    def update_progress(self, current: int, total: int, file_size_gb: float):
        """Update progress bar."""
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = current

        partial_gb = (current / total) * file_size_gb
        status_text = (
            f"Extracting file {self.extractor.status.current_file_index} of "
            f"{self.extractor.status.total_files} ({partial_gb:.2f} GB extracted "
            f"from {self.extractor.status.total_size_gb:.2f} GB total)"
        )
        self.status_label.config(text=status_text)
        self.root.update_idletasks()

    def start_extraction(self):
        """Start extraction process."""
        destination = self.destination_var.get()
        if not destination:
            messagebox.showwarning(
                "Warning", "Please select a destination folder.")
            return

        files = self.file_listbox.get(0, tk.END)
        if not files:
            self.show_countdown("No files selected for extraction.")
            return

        # Calculate total statistics
        self.extractor.status.total_files = len(files)
        self.extractor.status.total_size_gb = sum(
            os.path.getsize(file) / (1024 ** 3) for file in files
        )

        # Clear previous results
        self.results_text.delete(1.0, tk.END)

        # Run in separate thread
        thread = threading.Thread(target=self.extract_files,
                                  args=(files, destination), daemon=True)
        thread.start()

    def extract_files(self, files: List[str], destination: str):
        """Execute file extraction."""
        try:
            for idx, file_path in enumerate(files, start=1):
                self.extractor.status.current_file_index = idx

                result = self.extractor.extract_archive(
                    file_path, destination, self.update_progress
                )

                self.results_text.insert(tk.END, result + "\n")
                self.results_text.see(tk.END)
                self.root.update_idletasks()

            # Successful completion
            final_text = (
                f"✅ Extraction completed: {self.extractor.status.total_files} "
                f"file(s) / {self.extractor.status.total_size_gb:.2f} GB"
            )
            self.status_label.config(text=final_text)
            self.show_countdown(
                "✅ Extraction completed. Window will close automatically.")

        except (IOError, OSError, zipfile.BadZipFile,
                rarfile.Error, py7zr.exceptions.Bad7zFile) as e:
            error_msg = f"❌ Error during extraction: {str(e)}"
            self.results_text.insert(tk.END, error_msg + "\n")
            self.status_label.config(text=error_msg)

        finally:
            self.progress_bar['value'] = 0

    def safe_update_countdown(self, seconds: int):
        """Safely update countdown label."""
        if (self.countdown_label and
            self.countdown_window and
                self.countdown_running):
            try:
                self.countdown_label.config(
                    text=f"Closing in {seconds} seconds…")
            except tk.TclError:
                # Window was closed, stop countdown
                self.countdown_running = False

    def show_countdown(self, message: str, seconds: int = 60):
        """Display countdown for closing."""
        def countdown():
            self.countdown_running = True
            for sec in range(seconds, 0, -1):
                if not self.countdown_running:
                    break
                # Use after to update UI in main thread
                self.root.after(0, lambda s=sec: self.safe_update_countdown(s))
                time.sleep(1)

            if self.countdown_running:
                self.root.after(0, self.root.quit)

        # Create warning window
        self.countdown_window = tk.Toplevel(self.root)
        self.countdown_window.title("Warning")
        self.countdown_window.geometry("400x120")
        self.countdown_window.resizable(False, False)
        self.countdown_window.protocol("WM_DELETE_WINDOW", self.stop_countdown)

        tk.Label(self.countdown_window, text=message,
                 font=("Arial", 12), wraplength=350).pack(pady=10)

        self.countdown_label = tk.Label(self.countdown_window,
                                        text=f"Closing in {seconds} seconds…",
                                        font=("Arial", 10))
        self.countdown_label.pack()

        # Button to cancel countdown
        tk.Button(self.countdown_window, text="Cancel",
                  command=self.stop_countdown).pack(pady=5)

        threading.Thread(target=countdown, daemon=True).start()

    def stop_countdown(self):
        """Stop the countdown."""
        self.countdown_running = False
        if self.countdown_window:
            self.countdown_window.destroy()
            self.countdown_window = None
            self.countdown_label = None


def main():
    """Main function."""
    root = tk.Tk()
    ExtractionGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
