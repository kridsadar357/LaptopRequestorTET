# TET Asset Control Outside

This project is a C# Windows application for controlling and activating laptops intended for use outside of a specific network environment. The application utilizes MaterialSkin for modern UI styling and manages activation based on Wi-Fi SSID and a code-based activation process.

## Features

- **Activation Form**: The application presents a secure activation form to users when their device is not connected to a specified Wi-Fi network.
- **SSID Monitoring**: It continuously monitors the Wi-Fi SSID to determine whether the laptop is connected to an allowed network.
- **Activation Code Verification**: Users must enter a valid activation code to activate their laptop for use outside the specified Wi-Fi network.
- **Form Lock**: The activation form is locked, preventing the user from moving or closing it until a valid code is provided.
- **MaterialSkin UI**: The UI is built using MaterialSkin to provide a modern and visually appealing user interface.

## Project Structure

- **MaterialFormWithNoMove**: Custom form class that inherits from MaterialForm. This form prevents the user from moving the window or accessing the system context menu.
- **GetCurrentSSID**: Retrieves the SSID of the currently connected Wi-Fi network to determine the allowed status.
- **CheckActivationStatus**: Checks the activation status by reading a token from a local file, ensuring the laptop is activated for use.
- **ShowInputForm**: Displays the activation form to the user, requiring them to input a 4-digit activation code. The form remains on screen until a valid code is entered.
- **MonitorWiFi**: Continuously monitors the connected Wi-Fi SSID to determine if the activation form should be shown or hidden.

## Dependencies

- **MaterialSkin**: Used for building a modern UI with Material Design components.
- **System.Net.Http**: Used for sending HTTP requests to verify activation codes.

## Usage

1. **Build and Run**: Compile the project and run the executable on a Windows machine.
2. **Activation Process**:
   - If the device is not connected to an allowed Wi-Fi network (e.g., "TAKANE_WiFi" or "TAKANE_WH"), the activation form will be displayed.
   - Users must input a 4-digit code, which is then verified via an HTTP POST request.
   - Upon successful activation, the activation state is saved locally, allowing the laptop to operate outside the specified network for a limited time.

## Important Code Highlights

- **WndProc Override**: The `MaterialFormWithNoMove` class overrides the `WndProc` method to disable form movement and context menu actions.
- **Activation Validation**: The activation code is validated through an API (`active_outside.php`) and if valid, the activation token is saved to the local system to permit temporary usage.
- **Persistent Monitoring**: The application runs a background thread to monitor the Wi-Fi SSID and take necessary actions based on the network status.

## Installation

1. **Clone Repository**: Clone the project repository.
2. **Build Application**: Open the project in Visual Studio and build it.
3. **Deployment**: Use a setup project or a deployment tool to install the application on the target machines.

## Configuration

- **Allowed SSIDs**: Update the `validSSIDs` list in the code to change the allowed Wi-Fi networks.
- **Activation Endpoint**: Update the endpoint URL in the `ShowInputForm` method to match the backend service you are using for activation.

## Security Considerations

- The activation token includes a timestamp and a secret key to ensure it is unique and can be verified by the backend.
- The application prevents unauthorized use by locking the form and ensuring users cannot close or move it without entering a valid activation code.

## License

This project is proprietary and should not be distributed or copied without permission from Takane Technologies.

## Contact

For support or inquiries, please contact the project maintainer at [kridsadar@takane-th.com](mailto:kridsadar@takane-th.com).
