
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
from pathlib import Path

from utils.logger import logger


class SessionPickerDialog:
    """
    A simple listbox dialog showing all sessions from the store.
    Calls on_select(session_id) when the user confirms.
    """

    def __init__(
        self,
        sessions: list[dict],
        on_select: Callable[[str], None],
        on_delete: Optional[Callable[[str], None]] = None,
    ):
        self._sessions = sessions
        self._on_select = on_select
        self._on_delete = on_delete
        self._selected_id: Optional[str] = None

    def show(self) -> None:
        root = tk.Tk()
        root.title("Select Session to Replay")
        root.geometry("640x420")
        root.resizable(True, True)

       
        tk.Label(
            root,
            text="Recorded Sessions",
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        ).pack(fill="x", padx=16, pady=(14, 4))

       
        frame = tk.Frame(root)
        frame.pack(fill="both", expand=True, padx=16, pady=4)

        cols = ("name", "created", "events", "status")
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        tree.heading("name",    text="Session name")
        tree.heading("created", text="Created")
        tree.heading("events",  text="Events")
        tree.heading("status",  text="Status")
        tree.column("name",    width=240)
        tree.column("created", width=160)
        tree.column("events",  width=70,  anchor="center")
        tree.column("status",  width=90,  anchor="center")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for s in self._sessions:
            created = s.get("created_at", "")[:19].replace("T", " ")
            tree.insert("", "end", iid=s["id"], values=(
                s.get("name", "Untitled"),
                created,
                s.get("event_count", 0),
                s.get("status", ""),
            ))

        # Button row
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=16, pady=10)

        def on_replay():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("No selection", "Please select a session.")
                return
            self._selected_id = sel[0]
            root.destroy()
            self._on_select(self._selected_id)

        def on_delete():
            sel = tree.selection()
            if not sel:
                return
            sid = sel[0]
            name = tree.item(sid, "values")[0]
            if messagebox.askyesno("Delete session", f"Delete '{name}'?"):
                tree.delete(sid)
                if self._on_delete:
                    self._on_delete(sid)

        tk.Button(
            btn_frame, text="Replay selected",
            command=on_replay,
            bg="#1D9E75", fg="white",
            font=("Segoe UI", 10),
            padx=14, pady=6,
        ).pack(side="left", padx=(0, 8))

        if self._on_delete:
            tk.Button(
                btn_frame, text="Delete",
                command=on_delete,
                fg="#993C1D",
                font=("Segoe UI", 10),
                padx=14, pady=6,
            ).pack(side="left")

        tk.Button(
            btn_frame, text="Cancel",
            command=root.destroy,
            font=("Segoe UI", 10),
            padx=14, pady=6,
        ).pack(side="right")

        root.mainloop()


class RecordingNameDialog:
    def __init__(self, default: str = "Untitled session"):
        self._default = default
        self.result: Optional[str] = None

    def show(self) -> Optional[str]:
        root = tk.Tk()
        root.title("New Recording")
        root.geometry("420x160")
        root.resizable(False, False)

        tk.Label(
            root,
            text="Session name:",
            font=("Segoe UI", 10),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(20, 4))

        var = tk.StringVar(value=self._default)
        entry = tk.Entry(root, textvariable=var, font=("Segoe UI", 11), width=40)
        entry.pack(fill="x", padx=20)
        entry.select_range(0, "end")
        entry.focus()

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=20, pady=16)

        def confirm(event=None):
            self.result = var.get().strip() or self._default
            root.destroy()

        def cancel():
            self.result = None
            root.destroy()

        tk.Button(
            btn_frame, text="Start Recording",
            command=confirm,
            bg="#185FA5", fg="white",
            font=("Segoe UI", 10),
            padx=12, pady=5,
        ).pack(side="left")

        tk.Button(
            btn_frame, text="Cancel",
            command=cancel,
            font=("Segoe UI", 10),
            padx=12, pady=5,
        ).pack(side="right")

        root.bind("<Return>", confirm)
        root.bind("<Escape>", lambda _: cancel())
        root.mainloop()
        return self.result
