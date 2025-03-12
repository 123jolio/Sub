from PIL import Image
import os

def gif_to_frames(gif_path):
    # Get the directory where the Python script exists (current directory)
    current_dir = os.getcwd()

    # Load the GIF
    with Image.open(gif_path) as img:
        i = 0
        while True:
            try:
                img.seek(i)  # Move to the next frame
                frame = img.convert('RGB')  # Convert the frame to RGB
                # Save the frame with the name frame_0.jpg, frame_1.jpg, etc.
                frame.save(os.path.join(current_dir, f'frame_{i}.jpg'), 'JPEG')
                i += 1
            except EOFError:
                break  # No more frames

    print("GIF has been processed. Frames saved as JPGs in:", current_dir)

if __name__ == '__main__':
    # Get the current directory
    current_dir = os.getcwd()

    # Find the first GIF file in the current directory
    gif_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.gif')]

    if gif_files:
        gif_path = os.path.join(current_dir, gif_files[0])
        gif_to_frames(gif_path)
    else:
        print("No GIF file found in the current directory.")
