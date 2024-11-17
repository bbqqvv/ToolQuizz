import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import random
import time
import csv


class QuizApp:
    def __init__(self, root, question_data):
        self.root = root
        self.root.title("Quiz Application")
        self.root.geometry("700x700")
        self.root.config(bg="#ffffff")

        self.question_data = question_data
        self.current_question = 0
        self.score = 0
        self.selected_option = tk.StringVar()
        self.total_questions = 20

        self.answer_times = []

        # Timer setup
        self.timer = 1200
        self.timer_label = tk.Label(self.root, text=f"Time left: {self.format_time(self.timer)}",
                                    font=("Arial", 14, "bold"), fg="#ff5252", bg="#ffffff")
        self.timer_label.pack(pady=10)

        self.selected_questions = random.sample(self.question_data, min(self.total_questions, len(self.question_data)))

        self.create_ui()
        self.countdown()

    def format_time(self, seconds):
        """Format time in MM:SS."""
        minutes = seconds // 60
        seconds = seconds % 60
        return f"{minutes:02}:{seconds:02}"

    def countdown(self):
        """Countdown timer."""
        if self.timer > 0:
            self.timer -= 1
            self.timer_label.config(text=f"Time left: {self.format_time(self.timer)}")
            self.root.after(1000, self.countdown)
        else:
            self.submit_answer(force=True)

    def create_ui(self):
        """Set up the main user interface."""
        self.content_frame = tk.Frame(self.root, bg="#ffffff", padx=20, pady=20)
        self.content_frame.pack(fill="both", expand=True)

        # Scrollable question frame
        self.canvas = tk.Canvas(self.content_frame, bg="#ffffff", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.content_frame, orient="vertical", command=self.canvas.yview)
        self.question_scroll_frame = tk.Frame(self.canvas, bg="#ffffff")

        self.question_scroll_frame.bind("<Configure>",
                                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.question_scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.question_label = tk.Label(self.question_scroll_frame, text="", wraplength=650,
                                       font=("Arial", 16, "bold"), justify="center", bg="#ffffff", fg="#333333")
        self.question_label.pack(pady=20)

        self.status_label = tk.Label(self.content_frame, text="Question 1/20",
                                     font=("Arial", 12), bg="#ffffff", fg="#666666")
        self.status_label.pack(pady=10)

        self.progress = tk.IntVar()
        self.progress_bar = ttk.Progressbar(self.content_frame, length=500, variable=self.progress,
                                            maximum=self.total_questions, style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(pady=15)

        # Options frame
        self.options_frame = tk.Frame(self.content_frame, bg="#ffffff")
        self.options_frame.pack(pady=20)

        self.radio_buttons = []
        for i in range(4):
            rb = tk.Radiobutton(self.options_frame, text="", variable=self.selected_option, value=i,
                                font=("Arial", 12), bg="#ffffff", fg="#333333",
                                anchor="w", indicatoron=False, width=50, pady=5,
                                relief="ridge", bd=1)
            rb.pack(anchor="w", padx=10, pady=5)
            self.radio_buttons.append(rb)

        # Button frame
        self.button_frame = tk.Frame(self.content_frame, bg="#ffffff")
        self.button_frame.pack(pady=20)

        self.submit_button = tk.Button(self.button_frame, text="Submit", command=self.submit_answer,
                                       font=("Arial", 14), bg="#4CAF50", fg="white",
                                       activebackground="#45a049", activeforeground="white", width=15, bd=0)
        self.submit_button.pack(side="left", padx=10)

        self.restart_button = tk.Button(self.button_frame, text="Restart", command=self.restart_quiz,
                                        font=("Arial", 14), bg="#f44336", fg="white",
                                        activebackground="#e53935", activeforeground="white", width=15, bd=0)
        self.restart_button.pack(side="right", padx=10)
        self.restart_button.pack_forget()

        self.details_button = tk.Button(self.button_frame, text="Answer Details", command=self.show_answer_details,
                                        font=("Arial", 14), bg="#2196F3", fg="white",
                                        activebackground="#1976D2", activeforeground="white", width=15, bd=0)
        self.details_button.pack(side="right", padx=10)
        self.details_button.pack_forget()

        self.hint_button = tk.Button(self.button_frame, text="Hint", command=self.show_hint,
                                     font=("Arial", 14), bg="#FF9800", fg="white",
                                     activebackground="#F57C00", activeforeground="white", width=15, bd=0)
        self.hint_button.pack(side="left", padx=10)

        self.show_question()

    def show_question(self):
        """Display the current question and options."""
        question = self.selected_questions[self.current_question]
        self.question_label.config(text=f"Q{self.current_question + 1}: {question['Question']}")
        self.selected_option.set(-1)

        options = [question["Answer Option A"], question["Answer Option B"],
                   question["Answer Option C"], question["Answer Option D"]]
        for i, rb in enumerate(self.radio_buttons):
            rb.config(text=options[i])

        self.status_label.config(text=f"Question {self.current_question + 1}/{self.total_questions}")
        self.progress.set(self.current_question + 1)

        self.start_time = time.time()

    def submit_answer(self, force=False):
        """Submit the selected answer."""
        if not force and self.selected_option.get() == "-1":
            messagebox.showwarning("Warning", "Please select an answer!")
            return

        selected_index = int(self.selected_option.get()) if not force else -1
        question = self.selected_questions[self.current_question]

        correct_answers = question["Answer"].split(',')
        correct_answers = [answer.strip() for answer in correct_answers]

        user_answer = ["A", "B", "C", "D"][selected_index] if selected_index != -1 else None

        if user_answer in correct_answers:
            self.score += 1

        answer_time = time.time() - self.start_time
        self.answer_times.append({
            'question': question['Question'],
            'user_answer': user_answer if user_answer else "No answer",
            'correct_answer': ', '.join(correct_answers),
            'time_taken': round(answer_time, 2)
        })

        self.current_question += 1

        if self.current_question < len(self.selected_questions):
            self.show_question()
        else:
            self.show_result()

    def show_result(self):
        """Display the final result."""
        messagebox.showinfo("Quiz Completed", f"Your score is: {self.score}/{self.total_questions}")
        self.save_results()  # Save results to CSV
        self.submit_button.config(state="disabled")
        self.restart_button.pack(pady=10)
        self.details_button.pack(pady=10)

    def save_results(self):
        """Save quiz results to CSV."""
        filename = ".venv/quiz_results.csv"
        fieldnames = ['question', 'user_answer', 'correct_answer', 'time_taken']

        with open(filename, mode='a', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            if file.tell() == 0:  # Write header if file is empty
                writer.writeheader()

            for item in self.answer_times:
                writer.writerow(item)

        messagebox.showinfo("Saved", "Your quiz results have been saved!")

    def restart_quiz(self):
        """Restart the quiz."""
        self.current_question = 0
        self.score = 0
        self.timer = 1200
        self.answer_times = []
        self.selected_questions = random.sample(self.question_data, min(self.total_questions, len(self.question_data)))
        self.submit_button.config(state="normal")
        self.restart_button.pack_forget()
        self.details_button.pack_forget()
        self.show_question()
        self.countdown()

    def show_answer_details(self):
        """Show detailed answers."""
        result_window = tk.Toplevel(self.root)
        result_window.title("Answer Details")
        result_window.geometry("700x400")

        canvas = tk.Canvas(result_window)
        scrollbar = ttk.Scrollbar(result_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for item in self.answer_times:
            question_label = tk.Label(scrollable_frame, text=f"{item['question']}", font=("Arial", 12))
            user_answer_label = tk.Label(scrollable_frame, text=f"Your Answer: {item['user_answer']}",
                                         font=("Arial", 12))
            correct_answer_label = tk.Label(scrollable_frame, text=f"Correct Answer: {item['correct_answer']}",
                                            font=("Arial", 12))
            time_taken_label = tk.Label(scrollable_frame, text=f"Time Taken: {item['time_taken']}s", font=("Arial", 12))

            question_label.pack(pady=5)
            user_answer_label.pack(pady=5)
            correct_answer_label.pack(pady=5)
            time_taken_label.pack(pady=5)
            tk.Label(scrollable_frame, text="-" * 60).pack(pady=5)

        ok_button = tk.Button(result_window, text="OK", command=result_window.destroy, font=("Arial", 12),
                              bg="#4CAF50", fg="white")
        ok_button.pack(pady=10)

    def show_hint(self):
        """Show a hint for the current question."""
        question = self.selected_questions[self.current_question]
        hint = question.get("Hint", "No hint available.")
        messagebox.showinfo("Hint", hint)


def load_questions_from_excel(file_path):
    """Load questions from an Excel file."""
    try:
        df = pd.read_excel(file_path)
        print("Column names in the Excel file:", df.columns.tolist())
        df.columns = df.columns.str.strip()

        # Các cột cần thiết để quiz hoạt động
        required_columns = ['Question', 'Answer', 'Answer Option A', 'Answer Option B', 'Answer Option C',
                            'Answer Option D']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"Missing one or more required columns: {missing_columns}")

        # Đọc dữ liệu từ file Excel và chuyển đổi thành dictionary
        questions = df.to_dict(orient="records")

        # Duyệt qua các câu hỏi và gán "Answer" làm gợi ý (Hint)
        for question in questions:
            # Dùng cột 'Answer' làm gợi ý cho mỗi câu hỏi
            question['Hint'] = question['Answer']

        return questions

    except Exception as e:
        messagebox.showerror("Error", f"Error loading questions: {e}")
        return []


if __name__ == "__main__":
    excel_path = r"D:\StudyVKU\Year 3\Temp 1\CD 1\Java Web Advance-20241008T120028Z-001\Java Web Advance\JWEB\4_Quiz\JWEB_Question Bank_v1.1.xlsx"
    questions = load_questions_from_excel(excel_path)

    if questions:
        root = tk.Tk()
        app = QuizApp(root, questions)
        root.mainloop()