# example_usage.py - Complete example of how to use the User Interaction Replicator

import json
import os
import sys
from datetime import datetime
import pandas as pd

# Import our main classes
from user_interaction_replicator import UserInteractionReplicator
from desktop_automation import DesktopAutomationHandler

class InteractionReplicatorRunner:
    """
    Main runner class that orchestrates the entire replication process
    """
    
    def __init__(self, csv_file_path: str, config_file: str = "config.json"):
        self.csv_file_path = csv_file_path
        self.config = self._load_config(config_file)
        self.replicator = None
        self.desktop_handler = None
        
    def _load_config(self, config_file: str) -> dict:
        """Load configuration from JSON file"""
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    return json.load(f)
            else:
                # Return default configuration
                return {
                    "selenium": {"headless": False, "implicit_wait": 10},
                    "desktop_automation": {"failsafe": True, "pause": 0.1},
                    "timing": {"speed_multiplier": 1.0},
                    "logging": {"level": "INFO"}
                }
        except Exception as e:
            print(f"Error loading config: {e}")
            return {}
    
    def setup(self):
        """Initialize all components"""
        print("Setting up User Interaction Replicator...")
        
        # Initialize main replicator
        self.replicator = UserInteractionReplicator(
            csv_file_path=self.csv_file_path,
            headless=self.config.get("selenium", {}).get("headless", False)
        )
        
        # Initialize desktop automation handler
        self.desktop_handler = DesktopAutomationHandler(
            config=self.config.get("desktop_automation", {})
        )
        
        print("Setup completed!")
    
    def analyze_interactions(self):
        """Analyze the CSV data before replication"""
        print("\n=== INTERACTION ANALYSIS ===")
        
        # Basic statistics
        data = self.replicator.data
        print(f"Total interactions: {len(data)}")
        
        # Process breakdown
        if 'ProcessName' in data.columns:
            process_counts = data['ProcessName'].value_counts()
            print(f"\nApplications involved:")
            for process, count in process_counts.head(10).items():
                print(f"  {process}: {count} interactions")
        
        # Event type breakdown
        if 'Event' in data.columns:
            event_counts = data['Event'].value_counts()
            print(f"\nEvent types:")
            for event, count in event_counts.head(10).items():
                print(f"  {event}: {count} times")
        
        # Time span
        if 'StartTime' in data.columns and not data['StartTime'].isna().all():
            start_time = data['StartTime'].min()
            end_time = data['EndTime'].max() if 'EndTime' in data.columns else data['StartTime'].max()
            duration = end_time - start_time
            print(f"\nOriginal session duration: {duration}")
            print(f"Session started: {start_time}")
            print(f"Session ended: {end_time}")
        
        # Preview first few interactions
        print(f"\n=== PREVIEW OF FIRST 5 INTERACTIONS ===")
        preview = self.replicator.preview_interactions(5)
        print(preview.to_string())
    
    def run_filtered_replication(self, filters: dict = None):
        """Run replication with filters"""
        print(f"\n=== STARTING FILTERED REPLICATION ===")
        
        if filters:
            print(f"Applying filters: {filters}")
            filtered_data = self.replicator.filter_interactions(**filters)
            print(f"Filtered to {len(filtered_data)} interactions")
            
            # Temporarily replace data for this run
            original_data = self.replicator.data
            self.replicator.data = filtered_data.reset_index(drop=True)
        
        try:
            # Run replication
            results = self.replicator.replicate_interactions(
                speed_multiplier=self.config.get("timing", {}).get("speed_multiplier", 1.0)
            )
            
            # Print results
            self._print_results(results)
            
        finally:
            # Restore original data if it was filtered
            if filters:
                self.replicator.data = original_data
    
    def run_full_replication(self, start_index: int = 0, end_index: int = None):
        """Run full replication of all interactions"""
        print(f"\n=== STARTING FULL REPLICATION ===")
        
        if end_index is None:
            end_index = len(self.replicator.data)
        
        print(f"Replicating interactions {start_index} to {end_index}")
        
        results = self.replicator.replicate_interactions(
            start_index=start_index,
            end_index=end_index,
            speed_multiplier=self.config.get("timing", {}).get("speed_multiplier", 1.0)
        )
        
        self._print_results(results)
        
        return results
    
    def _print_results(self, results: dict):
        """Print replication results"""
        print(f"\n=== REPLICATION RESULTS ===")
        print(f"Total interactions processed: {results['total_interactions']}")
        print(f"Successful: {results['successful_interactions']}")
        print(f"Failed: {results['failed_interactions']}")
        print(f"Success rate: {(results['successful_interactions']/max(results['total_interactions'], 1))*100:.1f}%")
        print(f"Applications launched: {', '.join(results['applications_launched'])}")
        
        if results['errors']:
            print(f"\nErrors encountered:")
            for error in results['errors']:
                print(f"  - {error}")
    
    def run_interactive_mode(self):
        """Run in interactive mode for testing and debugging"""
        print("\n=== INTERACTIVE MODE ===")
        print("Commands:")
        print("  'analyze' - Analyze the CSV data")
        print("  'preview N' - Preview first N interactions")
        print("  'filter' - Apply filters and run replication")
        print("  'run N M' - Run interactions from index N to M")
        print("  'run' - Run all interactions")
        print("  'quit' - Exit")
        
        while True:
            try:
                command = input("\nEnter command: ").strip().lower()
                
                if command == 'quit':
                    break
                elif command == 'analyze':
                    self.analyze_interactions()
                elif command.startswith('preview'):
                    parts = command.split()
                    n = int(parts[1]) if len(parts) > 1 else 10
                    preview = self.replicator.preview_interactions(n)
                    print(preview.to_string())
                elif command == 'filter':
                    self._interactive_filter()
                elif command.startswith('run'):
                    parts = command.split()
                    if len(parts) == 1:
                        self.run_full_replication()
                    elif len(parts) == 3:
                        start_idx = int(parts[1])
                        end_idx = int(parts[2])
                        self.run_full_replication(start_idx, end_idx)
                    else:
                        print("Usage: 'run' or 'run start_index end_index'")
                else:
                    print("Unknown command")
                    
            except KeyboardInterrupt:
                print("\nInterrupted by user")
                break
            except Exception as e:
                print(f"Error: {e}")
    
    def _interactive_filter(self):
        """Interactive filter setup"""
        filters = {}
        
        process_name = input("Filter by process name (or Enter to skip): ").strip()
        if process_name:
            filters['process_name'] = process_name
        
        application = input("Filter by application (or Enter to skip): ").strip()
        if application:
            filters['application'] = application
        
        event_type = input("Filter by event type (or Enter to skip): ").strip()
        if event_type:
            filters['event_type'] = event_type
        
        if filters:
            self.run_filtered_replication(filters)
        else:
            print("No filters applied")


def main():
    """Main function to run the interaction replicator"""
    
    # Check command line arguments
    if len(sys.argv) < 2:
        print("Usage: python example_usage.py <csv_file_path> [config_file]")
        print("\nExample:")
        print("  python example_usage.py user_interactions.csv")
        print("  python example_usage.py user_interactions.csv my_config.json")
        return
    
    csv_file = sys.argv[1]
    config_file = sys.argv[2] if len(sys.argv) > 2 else "config.json"
    
    # Validate CSV file exists
    if not os.path.exists(csv_file):
        print(f"Error: CSV file '{csv_file}' not found!")
        return
    
    # Create runner and setup
    runner = InteractionReplicatorRunner(csv_file, config_file)
    runner.setup()
    
    # Analyze the data first
    runner.analyze_interactions()
    
    # Ask user what they want to do
    print(f"\n=== REPLICATION OPTIONS ===")
    print("1. Run interactive mode (recommended for first time)")
    print("2. Run full replication")
    print("3. Run filtered replication (web only)")
    print("4. Run first 10 interactions (test)")
    print("5. Exit")
    
    try:
        choice = input("\nSelect option (1-5): ").strip()
        
        if choice == '1':
            runner.run_interactive_mode()
        elif choice == '2':
            runner.run_full_replication()
        elif choice == '3':
            runner.run_filtered_replication({'process_name': 'chrome'})
        elif choice == '4':
            runner.run_full_replication(0, 10)
        elif choice == '5':
            print("Exiting...")
        else:
            print("Invalid choice")
            
    except KeyboardInterrupt:
        print("\nExiting...")
    except Exception as e:
        print(f"Error: {e}")


# Example configuration for different scenarios
def create_sample_configs():
    """Create sample configuration files for different use cases"""
    
    # Configuration for web automation only
    web_config = {
        "selenium": {
            "implicit_wait": 10,
            "page_load_timeout": 30,
            "headless": False,
            "browser_options": {
                "chrome": [
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-blink-features=AutomationControlled"
                ]
            }
        },
        "timing": {
            "speed_multiplier": 0.5,  # Slower for debugging
            "default_delay": 1.0,
            "max_delay": 3.0
        },
        "logging": {
            "level": "DEBUG",
            "file": "web_automation.log"
        }
    }
    
    # Configuration for desktop automation
    desktop_config = {
        "desktop_automation": {
            "failsafe": True,
            "pause": 0.2,
            "screenshot_on_error": True,
            "confidence": 0.8
        },
        "timing": {
            "speed_multiplier": 1.0,
            "default_delay": 0.5
        },
        "logging": {
            "level": "INFO",
            "file": "desktop_automation.log"
        }
    }
    
    # Production configuration
    production_config = {
        "selenium": {
            "implicit_wait": 5,
            "headless": True,  # Run in background
            "browser_options": {
                "chrome": [
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-blink-features=AutomationControlled",
                    "--window-size=1920,1080"
                ]
            }
        },
        "desktop_automation": {
            "failsafe": False,  # Disable for unattended operation
            "pause": 0.1
        },
        "timing": {
            "speed_multiplier": 2.0,  # Faster execution
            "max_delay": 2.0
        },
        "error_handling": {
            "max_retries": 3,
            "continue_on_error": True,
            "screenshot_on_failure": True
        }
    }
    
    # Write configuration files
    configs = {
        "config_web.json": web_config,
        "config_desktop.json": desktop_config,
        "config_production.json": production_config
    }
    
    for filename, config in configs.items():
        with open(filename, 'w') as f:
            json.dump(config, f, indent=4)
        print(f"Created {filename}")


if __name__ == "__main__":
    # Uncomment the next line to create sample configuration files
    # create_sample_configs()
    
    main()
