import uvicorn
import shutil
import pathlib

if __name__ == "__main__":
    wef = pathlib.Path(r"C:\Users\chris\AppData\Local\Microsoft\Office\16.0\Wef")
    if wef.exists():
        for item in wef.iterdir():
            shutil.rmtree(item) if item.is_dir() else item.unlink()
        print(f"Cleared Office cache: {wef}")

    uvicorn.run("app.api:app", host="0.0.0.0", port=8000, reload=False)
