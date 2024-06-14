// Initialize Microsoft Teams SDK
microsoftTeams.initialize();

var isPointerControlEnabled = false;

// Function to initialize pointer tool
function initializePointerTool() {
    // Add event listener for screen sharing
    microsoftTeams.videoConference.registerForVideoConferenceChange(onVideoConferenceChange);
}

// Function to handle video conference change event
function onVideoConferenceChange(change) {
    if (change.state === "Active" && change.action === "Start") {
        // Start screen sharing
        startScreenSharing();
    } else if (change.state === "Inactive" && change.action === "End") {
        // End screen sharing
        endScreenSharing();
    }
}

// Function to start screen sharing
function startScreenSharing() {
    // Add pointer UI to shared screen
    addPointerUI();
    // Add event listener for pointer movement
    microsoftTeams.videoConference.registerForPointerMove(onPointerMove);
}

// Function to end screen sharing
function endScreenSharing() {
    // Remove pointer UI from shared screen
    removePointerUI();
    // Remove event listener for pointer movement
    microsoftTeams.videoConference.unregisterForPointerMove();
}

// Function to add pointer UI to shared screen
function addPointerUI() {
    // Create pointer element
    var pointerElement = document.createElement("div");
    pointerElement.id = "pointer";
    pointerElement.style.position = "absolute";
    pointerElement.style.width = "10px";
    pointerElement.style.height = "10px";
    pointerElement.style.backgroundColor = "red";
    pointerElement.style.borderRadius = "50%";
    pointerElement.style.zIndex = "9999";
    document.body.appendChild(pointerElement);

    // Add event listener for pointer control
    pointerElement.addEventListener("mousedown", onPointerControlStart);
}

// Function to remove pointer UI from shared screen
function removePointerUI() {
    var pointerElement = document.getElementById("pointer");
    if (pointerElement) {
        pointerElement.parentNode.removeChild(pointerElement);
    }
}

// Function to handle pointer movement
function onPointerMove(pointerPosition) {
    var pointerElement = document.getElementById("pointer");
    if (pointerElement) {
        pointerElement.style.left = pointerPosition.x + "px";
        pointerElement.style.top = pointerPosition.y + "px";
    }
}

// Function to handle pointer control start
function onPointerControlStart(event) {
    // Check if pointer control is enabled
    if (!isPointerControlEnabled) return;

    // Get pointer element
    var pointerElement = document.getElementById("pointer");

    // Calculate pointer offset
    var offsetX = event.clientX - pointerElement.offsetLeft;
    var offsetY = event.clientY - pointerElement.offsetTop;

    // Add event listeners for pointer control
    document.addEventListener("mousemove", onPointerControlMove);
    document.addEventListener("mouseup", onPointerControlEnd);

    // Function to handle pointer control move
    function onPointerControlMove(event) {
        pointerElement.style.left = event.clientX - offsetX + "px";
        pointerElement.style.top = event.clientY - offsetY + "px";
    }

    // Function to handle pointer control end
    function onPointerControlEnd() {
        // Remove event listeners for pointer control
        document.removeEventListener("mousemove", onPointerControlMove);
        document.removeEventListener("mouseup", onPointerControlEnd);
    }
}

// Function to enable/disable pointer control
function togglePointerControl() {
    isPointerControlEnabled = !isPointerControlEnabled;
}

// Initialize pointer tool when page loads
initializePointerTool();
