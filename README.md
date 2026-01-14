# 📄 The "Exorcist" for Empty Pages

> *"I have deleted you. I have backspaced you. I have highlighted you and pressed every button on my keyboard. WHY DO YOU EXIST?"* — A broken soul, 2 AM.

Congratulations. You have found the only piece of software that understands your rage.

You have a document. It is 10 pages long. Word says it is 11 pages long. The 11th page is a ghost. A void. A white abyss mocking your very existence. You try to delete it, and Word laughs, formatting your headers into Comic Sans just to spite you.

**Enter: The Exorcist.**

This script does not "edit" your file. It does not "nicely ask" Word to remove the page.
It reaches into the raw XML guts of your `.docx` file, rips out the empty page's soul, and stitches the body back together before Word even notices what happened.

## 🌈 "Features"

*   **It Works**: Unlike the "Delete" key on your keyboard, evidently.
*   **No Bloat**: Written in pure Python because I refuse to install a 500MB library just to delete a page break.
*   **Safety Backup**: It creates a regular `.backup.docx` just in case you somehow manage to break it. (We know it's you, not the code).
*   **The "Scorched Earth" Option**: Delete all empty pages at once. Because who needs to check? Live dangerously.

## 🛠️ Usage Instructions (For Dummies)

1.  **Get the script**. If you can't figure this part out, I can't help you.
2.  **Run the script**:
    ```bash
    python3 find_empty_pages.py
    ```
3.  **Type the file name**. Accuracy helps.
4.  **Press buttons**. 'a' for "Destroy everything", 's' for "Micromanage".

## 🐛 Troubleshooting

*   **"It didn't work!"**: You probably typed the filename wrong.
*   **"My document is corrupted!"**: That's why I made a backup, genius. Restore it.
*   **"Word is still showing a blank page"**: Check your printer settings. Or your eyes. Or maybe the page just hates you personally.

## 🤝 Contributing
Don't. It works. Go away.

---

