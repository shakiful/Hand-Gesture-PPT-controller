# 🖐️ Hand Gesture PowerPoint Controller

Control your PowerPoint presentations touch-free using hand gestures and your webcam.

This project uses **Computer Vision** (OpenCV & MediaPipe) to track your hand landmarks in real-time and translates them into PowerPoint commands via the Windows COM interface. It includes features for slide navigation and a dynamic "depth-based" zoom.

## Demo

Download the Media File to see the Demo

## ✨ Features

* **Touchless Navigation:** Swipe through slides using a simple pointer gesture.
* **Minority Report Style Zoom:** Pinch and move your hand physically closer to the camera to zoom in, and pull back to zoom out.
* **Instant Reset:** Make a fist to reset the slide view to normal.
* **Visual Feedback:** On-screen overlay shows the current detected gesture and command status.
* **Focus Management:** Automatically attempts to focus the PowerPoint window to ensure commands are received.

## 🛠️ Requirements

* **OS:** Windows 10/11 (Required for `pywin32` COM integration)
* **Software:** Microsoft PowerPoint
* **Hardware:** Standard Webcam

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git)
    cd YOUR_REPO_NAME
    ```

2.  **Install the dependencies:**
    ```bash
    pip install opencv-python mediapipe numpy pywin32 pyautogui
    ```

## 🚀 How to Run

1.  Open your PowerPoint presentation.
2.  Run the Python script:
    ```bash
    python main.py
    ```
    *(Note: Replace `main.py` with whatever you named your file)*
3.  The webcam window will open.
4.  **Start your Slideshow (F5)** in PowerPoint.
5.  Stand back so the camera can see your hand and start presenting!

## ✌️ Gesture Guide

| Action | Gesture Description | Visual Cue |
| :--- | :--- | :--- |
| **Next Slide** | **"Pointer Mode"**: Extend your Index finger (☝️) and move it to the **Right**. | A yellow trail line follows your finger. |
| **Previous Slide** | **"Pointer Mode"**: Extend your Index finger (☝️) and move it to the **Left**. | A yellow trail line follows your finger. |
| **Zoom In** | **"Pinch"**: Pinch Index & Thumb (👌). Move hand **Closer** to camera. | Status bar says "Zooming In". |
| **Zoom Out** | **"Pinch"**: Pinch Index & Thumb (👌). Move hand **Away** from camera. | Status bar says "Zooming Out". |
| **Reset View** | **"Fist"**: Clench your hand into a fist (✊). | Zoom resets to "Fit to Window". |

## ⚠️ Troubleshooting / Notes

* **Window Focus:** If the slides are not changing, click once on the PowerPoint window to ensure it is the "Active" window. The script tries to force focus, but Windows security sometimes blocks this.
* **Lighting:** Ensure your hand is well-lit. Strong backlighting (light behind you) makes it hard for MediaPipe to track fingers.
* **Zoom in Slideshow:** The script uses a `Ctrl + Scroll` simulation for zooming in Slideshow mode. This requires the PPT window to be active.

## 📜 License

This project is open-source and available under the [MIT License](LICENSE).
