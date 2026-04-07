# A 3D Raycasting Engine in Excel
![Ray](https://github.com/user-attachments/assets/93a9d5d4-a987-4be4-aec4-47404612d68d)

A high-performance 3D rendering engine built from scratch within Microsoft Excel. Leveraging optimized VBA-driven ray-marching logic, this project transforms a standard spreadsheet grid into a real-time, interactive environment featuring dynamic spatial calculations and first-person navigation.

## 🎮 Features

* **Dynamic Raycasting:** Real-time 3D rendering with distance-based shading.
* **Multi-Enemy AI:** An array-based spawn system with randomized placement and individual respawn timers.
* **Combat System:** Projectiles that scale with distance.
* **Live HUD:** Real-time tracking of Player Health, Kill Count, and Directional Minimap.
* **Tuning Menu:** In-sheet controls for Player Speed, Rotation Sensitivity, and Spawn Rates—no code changes required for gameplay balancing.
* **Directional Minimap:** A 2D top-down view with a directional arrow indicating player orientation.

## 🛠️ Technical Overview

As an engine built in a non-native environment, this project pushes the boundaries of VBA's execution speed. Key technical implementations include:

* **Linear Ray-Tracing Logic:** An optimised grid-traversal algorithm that calculates precise wall intersections in real-time, simulating 3D depth without native 3D libraries.
* **Array Rendering:** To avoid the overhead of individual cell updates, the engine writes to a 2D array and flushes to the sheet in a single operation.
* **Windows API Integration:** Utilises low-level system hooks to enable simultaneous movement and rotation, creating a fluid navigation experience
* **Perspective Scaling:** 3D sprite projection for enemies and projectiles using vertical and horizontal scaling.

## 🕹️ Controls

| Key | Action |
| :--- | :--- |
| **W / S** | Move Forward / Backward |
| **A / D** | Rotate Camera Left / Right |
| **Space** | Fire Plasma Bolt |
| **Esc** | Quit Macro |

## 📦 Installation & Setup

1.  **Download:** Clone this repository or download the `.xlsm` file to a dedicated folder on your local drive.
2.  **Unblock the File:** Before opening, right-click the `.xlsm` file > **Properties** > check the **Unblock** box at the bottom > **OK**. (This is a standard Windows security requirement for macro-enabled files).
3.  **Set Trusted Location (Recommended):** To bypass security pop-ups permanently for this project:
    * Open Excel and go to **File > Options > Trust Center**.
    * Click **Trust Center Settings...** > **Trusted Locations**.
    * Click **Add new location...**, browse to your project folder, and click **OK**.
4.  **Open the Source Code:** * Launch the file and press `Alt + F11` to open the **VBA Editor**.
    * In the "Project" pane on the left, double-click the module containing the engine code (e.g., `Module1`).
5.  **Run the Engine:** * Click anywhere inside the `StartRaycaster` subroutine.
    * Press `F5` to initialize the engine loop.
6.  **Maintain Input Focus:** * **Crucial:** After pressing `F5`, click back into the **VBA Editor window**. 
    * By keeping the VBA window as the active focus, Excel will not attempt to "type" into the spreadsheet cells, allowing the engine to capture your WASD inputs cleanly.

---

*Battlefield on a Spreadsheet*
