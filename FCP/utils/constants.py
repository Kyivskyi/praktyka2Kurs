CONVERTIBLE_FORMATS = {
    ".docx": [".pdf"],
    ".pdf": [".txt", ".docx"],
    ".txt": [".pdf", ".docx"],
    ".jpg": [".png", ".webp"],
    ".png": [".jpg", ".webp"],
    ".csv": [".xlsx", ".json"],
    ".json": [".csv", ".xlsx"],
    ".mp4": [".mkv", ".mov", ".flv", ".wmv", ".mp3"],
    ".mkv": [".mp4", ".avi", ".mov", ".flv", ".wmv"],
    ".mov": [".mp4", ".avi", ".mkv", ".flv", ".wmv"],
    ".flv": [".mp4", ".avi", ".mkv", ".mov", ".wmv"],
    ".wmv": [".mp4", ".avi", ".mkv", ".mov", ".flv"],
    ".mp3": [".wav", ".ogg"],
    ".wav": [".mp3", ".ogg"],
    ".ogg": [".mp3", ".wav"]
}