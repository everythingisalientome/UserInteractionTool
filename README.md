# User Interaction Replication Tool

A comprehensive Python tool that reads CSV files containing captured user interactions and replicates them using Selenium WebDriver for web applications and desktop automation libraries for native applications.

## Features

- **Web Application Automation**: Supports Chrome, Firefox, Edge, and other web browsers
- **Desktop Application Automation**: Handles Windows desktop applications including Office suite
- **Flexible Configuration**: JSON-based configuration for different scenarios
- **Interactive Mode**: Test and debug interactions step by step
- **Comprehensive Logging**: Detailed logs for troubleshooting
- **Filter Support**: Filter interactions by application, event type, etc.
- **Error Handling**: Robust error handling with retry mechanisms

## Prerequisites

### System Requirements
- Python 3.8 or higher
- Windows 10/11 (for desktop automation features)
- Chrome browser (for web automation)

### Required Python Packages
```bash
pip install -r requirements.txt
```

### Additional Setup
1. **ChromeDriver**: Will be automatically managed by `webdriver-manager`
2. **For desktop automation**: Install optional Windows-specific packages:
   ```bash
   pip install pywin32 opencv-python pillow
   ```

## Installation

1. **Clone or download the project files**
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Create configuration file** (optional - defaults will be used):
   ```bash
   python example_usage.py --create-configs
   ```

## CSV File Format

Your CSV file should contain the following columns (as mentioned in your original request):

| Column Name | Description |
|-------------|-------------|
| ProcessName | Name of the application process |
| UserName | User who performed the action |
| ExeName | Executable name |
| Application | Application name |
| Event | Type of event (click, type, etc.) |
| FieldName | Name/ID of the UI element |
| FieldType | Type of field (button, text, etc.) |
| Sentence | Text content to input |
| URL | Web page URL (for web interactions) |
| StartTime | When the interaction started |
| EndTime | When the interaction ended |
| WindowName | Window title |
| ... | (and other columns as specified) |

## Quick Start

### 1. Basic Usage
```bash
python example_usage.py your_interactions.csv
```

### 2. Interactive Mode (Recommended for first time)
```python
from example_usage import InteractionReplicatorRunner

runner = InteractionReplicatorRunner("your_interactions.csv")
runner.setup()
runner.run_interactive_mode()
```

### 3. Programmatic Usage
```python
from user_interaction_replicator import UserInteractionReplicator

# Initialize the replicator
replicator = UserInteractionReplicator("your_interactions.csv")

# Preview interactions
print(replicator.preview_interactions(10))

# Run replication
results = replicator.replicate_interactions(
    start_index=0,
    end_index=50,
    speed_multiplier=0.5  # Run at half speed
)

print(f"Success rate: {results['successful_interactions']/results['total_interactions']*100:.1f}%")
```

## Configuration

The tool uses JSON configuration files. Here's an example:

```json
{
    "selenium": {
        "implicit_wait": 10,
        "page_load_timeout": 30,
        "headless": false,
        "browser_options": {
            "chrome": [
                "--no-sandbox",
                "--disable-dev-shm-usage"
            ]
        }
    },
    "desktop_automation": {
        "failsafe": true,
        "pause": 0.1,
        "screenshot_on_error": true,
        "confidence": 0.8
    },
    "timing": {
        "speed_multiplier": 1.0,
        "default_delay": 1.0,
        "max_delay": 5.0
    },
    "logging": {
        "level": "INFO",
        "file": "interaction_replication.log"
    }
}
```

### Configuration Options

#### Selenium Settings
- `headless`: Run browser without GUI
- `implicit_wait`: Time to wait for elements
- `page_load_timeout`: Maximum page load time
- `browser_options`: Additional Chrome options

#### Desktop Automation Settings
- `failsafe`: Enable PyAutoGUI failsafe (move mouse to corner to stop)
- `pause`: Delay between actions
- `screenshot_on_error`: Take screenshots when errors occur
- `confidence`: Image matching confidence level

#### Timing Settings
- `speed_multiplier`: Speed up (>1) or slow down (<1) execution
- `default_delay`: Default delay between actions
- `max_delay`: Maximum allowed delay

## Advanced Features

### 1. Filtering Interactions
```python
# Filter by application
web_interactions = replicator.filter_interactions(process_name="chrome")

# Filter by event type
click_events = replicator.filter_interactions(event_type="click")

# Multiple filters
filtered = replicator.filter_interactions(
