import cv2
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import socket
import pickle
import struct
import time
from PIL import Image, ImageTk
import pyvirtualcam
import win32com.client

# Constants for network discovery and control messages
DISCOVERY_PORT = 9998
CONTROL_PORT = 9997
START_MESSAGE = b"START_STREAMING"
STOP_MESSAGE = b"STOP_STREAMING"

class WebcamStreamer:
    def __init__(self, root, mode):
        self.root = root
        self.root.title("USB Webcam Streamer")
        self.mode = mode  # Mode is now set during initialization
        self.webcam_index = tk.IntVar(value=0)
        self.ip_address = tk.StringVar(value="127.0.0.1")
        self.port = tk.IntVar(value=9999)
        self.capture = None
        self.streaming = False
        self.broadcasting_enabled = False  # Start with broadcasting disabled
        self.control_listener_thread = None  # Keep track of control listener thread

        # Variables for bitrate calculation
        self.bytes_sent = 0
        self.bytes_received = 0
        self.last_update_time = time.time()

        self.setup_gui()

        # Start discovery listener on the client
        if self.mode == "Client":
            threading.Thread(target=self.start_discovery_listener, daemon=True).start()
            self.status_label.config(text="Searching for server...", foreground="orange")
        else:
            self.status_label.config(text="Waiting to start...", foreground="orange")

    def setup_gui(self):
        # Webcam selection
        cam_frame = ttk.LabelFrame(self.root, text="Webcam Selection")
        cam_frame.pack(padx=10, pady=5, fill="x")

        if self.mode == "Client":
            ttk.Label(cam_frame, text="Webcam:").pack(side="left", padx=5, pady=5)
            self.cam_dropdown = ttk.Combobox(cam_frame, state="readonly")
            self.cam_dropdown['values'] = self.get_available_cams()
            self.cam_dropdown.current(0)
            self.cam_dropdown.bind("<<ComboboxSelected>>", self.update_webcam_index)
            self.cam_dropdown.pack(side="left", padx=5, pady=5)
        else:
            ttk.Label(cam_frame, text="Server Mode - No Webcam Selection Required").pack(padx=5, pady=5)

        # IP and Port
        network_frame = ttk.LabelFrame(self.root, text="Network Settings")
        network_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(network_frame, text="IP Address:").pack(side="left", padx=5, pady=5)
        self.ip_entry = ttk.Entry(network_frame, textvariable=self.ip_address, state="readonly" if self.mode == "Server" else "normal")
        self.ip_entry.pack(side="left", padx=5, pady=5)
        ttk.Label(network_frame, text="Port:").pack(side="left", padx=5, pady=5)
        ttk.Entry(network_frame, textvariable=self.port, state="readonly" if self.mode == "Server" else "normal").pack(side="left", padx=5, pady=5)
        self.status_label = ttk.Label(network_frame, text="Initializing...", foreground="orange")
        self.status_label.pack(side="left", padx=5, pady=5)

        # Start/Stop Button
        if self.mode == "Server":
            self.start_button = ttk.Button(self.root, text="Start", command=self.toggle_streaming)
            self.start_button.pack(pady=5)
        else:
            self.start_button = None

        # Video display label
        self.video_label = ttk.Label(self.root)
        self.video_label.pack(padx=10, pady=10)

        # Bitrate label
        self.bitrate_label = ttk.Label(self.root, text="Bitrate: 0 Mbps")
        self.bitrate_label.pack()

        # Set initial IP address
        if self.mode == "Server":
            self.ip_address.set(self.get_local_ip())
        else:
            self.ip_address.set("127.0.0.1")

    def broadcast_server_presence(self):
        # Server periodically broadcasts its presence
        self.broadcasting_enabled = True
        broadcast_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        broadcast_socket.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        server_ip = self.get_local_ip().encode()
        while self.broadcasting_enabled:
            try:
                broadcast_socket.sendto(server_ip, ('<broadcast>', DISCOVERY_PORT))
                print("Server: Broadcasting presence")
                time.sleep(1)  # Broadcast every second
            except Exception as e:
                print(f"Server: Error broadcasting presence: {e}")
                break
        broadcast_socket.close()
        print("Server: Stopped broadcasting presence")

    def start_discovery_listener(self):
        # Client listens for server broadcasts
        while True:
            if self.streaming:
                time.sleep(1)  # Pause discovery when streaming
                continue
            discovery_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            discovery_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            discovery_socket.bind(('', DISCOVERY_PORT))
            discovery_socket.settimeout(5)
            print("Client: Listening for server broadcasts")
            self.status_label.config(text="Searching for server...", foreground="orange")
            while self.ip_address.get() == "127.0.0.1" and not self.streaming:
                try:
                    data, addr = discovery_socket.recvfrom(1024)
                    server_ip = data.decode()
                    if server_ip != self.get_local_ip():
                        # Update IP address field
                        self.ip_address.set(server_ip)
                        print(f"Client: Discovered server IP: {server_ip}")
                        # Send start command to server
                        threading.Thread(target=self.send_start_command, daemon=True).start()
                        break
                except socket.timeout:
                    print("Client: No broadcast received. Retrying...")
                except Exception as e:
                    print(f"Client: Discovery listener error: {e}")
                    break
            discovery_socket.close()
            time.sleep(1)  # Wait before restarting the discovery listener

    def send_start_command(self):
        try:
            self.control_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.control_socket.connect((self.ip_address.get(), CONTROL_PORT))
            self.control_socket.sendall(START_MESSAGE)
            self.status_label.config(text="Sent start command to server", foreground="orange")
            print("Client: Sent start command to server")
            # Wait for server's confirmation to start streaming
            response = self.control_socket.recv(1024)
            if response == START_MESSAGE:
                print("Client: Received start confirmation from server")
                self.status_label.config(text="Starting streaming...", foreground="orange")
                # Start a thread to listen for control messages from the server
                threading.Thread(target=self.listen_for_server_messages, daemon=True).start()
                self.toggle_streaming()
            else:
                print("Client: Unexpected response from server:", response)
        except Exception as e:
            print(f"Client: Error sending start command: {e}")
            self.status_label.config(text=f"Error: {e}", foreground="red")

    def listen_for_server_messages(self):
        try:
            while self.streaming:
                data = self.control_socket.recv(1024)
                if not data:
                    # Server closed the connection
                    print("Client: Control connection closed by server")
                    self.status_label.config(text="Server disconnected", foreground="red")
                    self.streaming = False
                    self.cleanup_resources()
                    # Reset IP address and restart discovery
                    self.ip_address.set("127.0.0.1")
                    threading.Thread(target=self.start_discovery_listener, daemon=True).start()
                    self.status_label.config(text="Stopped. Searching for server...", foreground="orange")
                    break
                elif data == STOP_MESSAGE:
                    print("Client: Received stop command from server")
                    self.status_label.config(text="Server requested to stop", foreground="orange")
                    self.toggle_streaming()
                    break
                else:
                    print("Client: Received unknown control message:", data)
        except Exception as e:
            print(f"Client: Error receiving control message: {e}")

    def start_control_listener(self):
        # Server listens for start/stop commands from the client
        self.control_listener_enabled = True
        control_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        control_socket.bind(('', CONTROL_PORT))
        control_socket.listen(1)
        print("Server: Control listener started")
        while self.control_listener_enabled:
            try:
                conn, addr = control_socket.accept()
                print(f"Server: Received control connection from {addr}")
                data = conn.recv(1024)
                if data == START_MESSAGE:
                    print("Server: Received start command from client")
                    self.status_label.config(text="Starting streaming...", foreground="orange")
                    if not self.streaming:
                        # Send confirmation back to client
                        conn.sendall(START_MESSAGE)
                        # Store the control connection
                        self.control_conn = conn
                        self.toggle_streaming()
                    else:
                        print("Server: Already streaming")
                        conn.sendall(b"ALREADY_STREAMING")
                elif data == STOP_MESSAGE:
                    print("Server: Received stop command from client")
                    self.status_label.config(text="Stopping streaming...", foreground="orange")
                    if self.streaming:
                        self.toggle_streaming()
                    conn.sendall(STOP_MESSAGE)
                    conn.close()
                else:
                    print("Server: Unknown control message:", data)
                    conn.close()
            except Exception as e:
                if self.control_listener_enabled:
                    print(f"Server: Control listener error: {e}")
                break
        control_socket.close()
        print("Server: Control listener stopped")

    def get_available_cams(self):
        # Get device names using WMI
        dev_enum = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        wbem_services = dev_enum.ConnectServer(".", "root\\cimv2")
        wbem_devices = wbem_services.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0")

        device_names = []
        index = 0
        for device in wbem_devices:
            if "USB" in device.PNPDeviceID and "vid" in device.PNPDeviceID.lower():
                name = device.Name
                # Check if the device can be opened by OpenCV
                cap = cv2.VideoCapture(index)
                if cap.isOpened():
                    device_names.append(f"{name} ({index})")
                    cap.release()
                index += 1

        if not device_names:
            device_names = ["Default Camera (0)"]
        return device_names

    def update_webcam_index(self, event):
        selected = event.widget.get()
        # Extract index from the device name
        index = int(selected.split('(')[-1].strip(')'))
        self.webcam_index.set(index)

    def toggle_streaming(self):
        if not self.streaming:
            if self.mode == "Server":
                self.start_button.config(text="Stop")
                # Start broadcasting presence
                threading.Thread(target=self.broadcast_server_presence, daemon=True).start()
                # Start control listener
                self.control_listener_thread = threading.Thread(target=self.start_control_listener, daemon=True)
                self.control_listener_thread.start()
                self.status_label.config(text="Waiting for client...", foreground="orange")
            self.streaming = True
            if self.mode == "Client":
                threading.Thread(target=self.start_client_streaming, daemon=True).start()
            else:
                # The server streaming will start after receiving the start command from the client
                pass
        else:
            if self.mode == "Server":
                self.start_button.config(text="Start")
                # Stop broadcasting
                self.broadcasting_enabled = False
                # Stop control listener
                self.control_listener_enabled = False
                # Send stop command to client
                if hasattr(self, 'control_conn'):
                    try:
                        self.control_conn.sendall(STOP_MESSAGE)
                        print("Server: Sent stop command to client")
                        self.control_conn.close()
                        del self.control_conn
                    except Exception as e:
                        print(f"Server: Error sending stop command to client: {e}")
                else:
                    print("Server: No control connection to client")
                self.status_label.config(text="Stopped. Waiting to start...", foreground="orange")
            self.streaming = False
            self.status_label.config(text="Disconnected", foreground="red")
            # Clean up resources
            self.cleanup_resources()
            if self.mode == "Client":
                # Reset IP address and restart discovery
                self.ip_address.set("127.0.0.1")
                threading.Thread(target=self.start_discovery_listener, daemon=True).start()
                self.status_label.config(text="Stopped. Searching for server...", foreground="orange")

    def cleanup_resources(self):
        # Close sockets and resources if they exist
        if hasattr(self, 'conn'):
            try:
                self.conn.close()
            except:
                pass
            del self.conn
        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except:
                pass
            del self.server_socket
        if hasattr(self, 'virtual_cam'):
            try:
                self.virtual_cam.close()
            except:
                pass
            del self.virtual_cam
        if hasattr(self, 'client_socket'):
            try:
                self.client_socket.close()
            except:
                pass
            del self.client_socket
        if hasattr(self, 'capture') and self.capture is not None:
            try:
                self.capture.release()
            except:
                pass
            del self.capture
        if hasattr(self, 'control_socket'):
            try:
                self.control_socket.close()
            except:
                pass
            del self.control_socket
        if hasattr(self, 'control_conn'):
            try:
                self.control_conn.close()
            except:
                pass
            del self.control_conn

    def start_client_streaming(self):
        try:
            self.status_label.config(text="Connecting to server...", foreground="orange")
            # Open the webcam using the selected index
            self.capture = cv2.VideoCapture(self.webcam_index.get())
            if not self.capture.isOpened():
                messagebox.showerror("Error", "Cannot open webcam")
                self.streaming = False
                self.status_label.config(text="Disconnected", foreground="red")
                return

            # Allow the camera to warm up
            time.sleep(1)

            # Initialize bytes sent for bitrate calculation
            self.bytes_sent = 0
            self.last_update_time = time.time()

            self.client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.client_socket.connect((self.ip_address.get(), self.port.get()))
            print(f"Client: Connected to server at {self.ip_address.get()}:{self.port.get()}")
            self.status_label.config(text="Connected", foreground="green")

            self.update_video_frame()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.streaming = False
            self.status_label.config(text="Disconnected", foreground="red")

    def update_video_frame(self):
        if self.streaming:
            ret, frame = self.capture.read()
            if ret:
                # Serialize frame
                data = pickle.dumps(frame)
                # Send message length first
                message_size = struct.pack("Q", len(data))
                try:
                    self.client_socket.sendall(message_size + data)
                    # Update bytes sent
                    self.bytes_sent += len(message_size + data)
                except Exception as e:
                    print(f"Client: Failed to send frame: {e}")
                    self.status_label.config(text="Disconnected", foreground="red")
                    self.streaming = False
                    self.cleanup_resources()
                    # Reset IP address and restart discovery
                    self.ip_address.set("127.0.0.1")
                    threading.Thread(target=self.start_discovery_listener, daemon=True).start()
                    self.status_label.config(text="Stopped. Searching for server...", foreground="orange")
                    return

                # Display the frame in the GUI
                cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(cv2image)
                imgtk = ImageTk.PhotoImage(image=img)
                self.video_label.imgtk = imgtk
                self.video_label.configure(image=imgtk)

                # Update bitrate every second
                current_time = time.time()
                if current_time - self.last_update_time >= 1:
                    bitrate = (self.bytes_sent * 8) / (current_time - self.last_update_time) / 1_000_000  # Mbps
                    self.bitrate_label.config(text=f"Bitrate: {bitrate:.2f} Mbps")
                    self.bytes_sent = 0
                    self.last_update_time = current_time

                # Schedule the next frame update
                self.root.after(10, self.update_video_frame)
            else:
                print("Client: Failed to grab frame")
                messagebox.showerror("Error", "Failed to grab frame from webcam")
                self.streaming = False
                self.cleanup_resources()
                self.status_label.config(text="Disconnected", foreground="red")
        else:
            self.cleanup_resources()
            self.status_label.config(text="Disconnected", foreground="red")

    def start_server_streaming(self):
        try:
            self.status_label.config(text="Waiting for connection...", foreground="orange")
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.bind(('', self.port.get()))  # Bind to all interfaces
            self.server_socket.listen(5)
            print(f"Server: Listening on {self.get_local_ip()}:{self.port.get()}")

            self.bytes_received = 0
            self.last_update_time = time.time()

            threading.Thread(target=self.accept_client_connection, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.streaming = False
            if self.start_button:
                self.start_button.config(text="Start")
            if hasattr(self, 'server_socket'):
                self.server_socket.close()
            self.status_label.config(text="Disconnected", foreground="red")

    def accept_client_connection(self):
        try:
            self.conn, addr = self.server_socket.accept()
            print(f"Server: Connection from {addr}")
            self.status_label.config(text="Connected", foreground="green")

            # Initialize data variables
            self.data = b""
            self.payload_size = struct.calcsize("Q")

            # Start receiving frames
            self.receive_frame()
        except Exception as e:
            print(f"Server: Error accepting connection: {e}")
            self.streaming = False
            self.status_label.config(text=f"Error: {e}", foreground="red")

    def receive_frame(self):
        if not self.streaming:
            self.cleanup_resources()
            self.status_label.config(text="Disconnected", foreground="red")
            return
        try:
            # Initialize virtual camera if not already
            if not hasattr(self, 'virtual_cam'):
                frame = self.get_next_frame()
                if frame is None:
                    return
                height, width, _ = frame.shape
                self.virtual_cam = pyvirtualcam.Camera(width=width, height=height, fps=20)
                print(f'Server: Virtual camera initialized: {self.virtual_cam.device}')
            else:
                frame = self.get_next_frame()
                if frame is None:
                    return

            # Display the frame in the GUI
            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(cv2image)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)

            # Send frame to virtual webcam
            self.virtual_cam.send(cv2image)
            self.virtual_cam.sleep_until_next_frame()

            # Update bytes received
            self.bytes_received += self.current_frame_size

            # Update bitrate every second
            current_time = time.time()
            if current_time - self.last_update_time >= 1:
                bitrate = (self.bytes_received * 8) / (current_time - self.last_update_time) / 1_000_000  # Mbps
                self.bitrate_label.config(text=f"Bitrate: {bitrate:.2f} Mbps")
                self.bytes_received = 0
                self.last_update_time = current_time

            # Schedule the next frame update
            self.root.after(10, self.receive_frame)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.streaming = False
            if self.start_button:
                self.start_button.config(text="Start")
            self.cleanup_resources()
            self.status_label.config(text="Disconnected", foreground="red")

    def get_next_frame(self):
        try:
            while len(self.data) < self.payload_size:
                packet = self.conn.recv(4 * 1024)  # 4K
                if not packet:
                    self.streaming = False
                    return None
                self.data += packet
            if not self.data:
                self.streaming = False
                return None
            packed_msg_size = self.data[:self.payload_size]
            self.data = self.data[self.payload_size:]
            msg_size = struct.unpack("Q", packed_msg_size)[0]
            self.current_frame_size = msg_size + self.payload_size  # For bitrate calculation

            while len(self.data) < msg_size:
                self.data += self.conn.recv(4 * 1024)
            frame_data = self.data[:msg_size]
            self.data = self.data[msg_size:]
            frame = pickle.loads(frame_data)
            return frame
        except Exception as e:
            print(f"Server: Error receiving frame: {e}")
            return None

    def get_local_ip(self):
        # Retrieve the local IP address
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            # Try to connect to an external IP address to get the local IP
            s.connect(('8.8.8.8', 80))
            local_ip = s.getsockname()[0]
        except Exception:
            # Fallback to '127.0.0.1'
            local_ip = '127.0.0.1'
        finally:
            s.close()
        return local_ip

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window until mode is selected

    # Create a startup popup to select mode using radio buttons
    mode_selection = tk.Toplevel(root)
    mode_selection.title("Select Mode")

    selected_mode = tk.StringVar(value="Server")

    def confirm_mode():
        mode_selection.destroy()
        root.deiconify()  # Show the main window
        app = WebcamStreamer(root, selected_mode.get())

    ttk.Label(mode_selection, text="Select Mode:").pack(padx=10, pady=5)
    ttk.Radiobutton(mode_selection, text="Server", variable=selected_mode, value="Server").pack(padx=10, pady=5)
    ttk.Radiobutton(mode_selection, text="Client", variable=selected_mode, value="Client").pack(padx=10, pady=5)
    ttk.Button(mode_selection, text="OK", command=confirm_mode).pack(pady=10)

    mode_selection.protocol("WM_DELETE_WINDOW", root.destroy)  # Ensure app exits if mode selection is closed
    root.mainloop()

if __name__ == "__main__":
    main()
