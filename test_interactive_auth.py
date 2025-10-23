#!/usr/bin/env python3
"""
Test script for interactive authentication
This script tests the interactive authentication without running the full application
"""

import sys
import os

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from Unidad_Accion.SharePointInteractiveAuth import SharePointInteractiveAuth
    print("✅ Successfully imported SharePointInteractiveAuth")
    
    # Test authentication
    print("\n🔐 Testing interactive authentication...")
    auth = SharePointInteractiveAuth()
    
    if auth.authenticate_interactive():
        print("✅ Interactive authentication successful!")
        
        # Test getting site ID
        print("\n🏢 Testing site ID retrieval...")
        site_id = auth.get_site_id()
        if site_id:
            print(f"✅ Site ID obtained: {site_id}")
        else:
            print("❌ Failed to get site ID")
        
        # Test getting available lists
        print("\n📋 Testing available lists...")
        lists = auth.get_available_lists()
        if lists:
            print(f"✅ Found {len(lists)} lists:")
            for lista in lists[:5]:  # Show first 5 lists
                print(f"  - {lista.get('displayName', 'Unknown')}")
        else:
            print("❌ Failed to get available lists")
        
        # Test getting data from Tutorias list
        print("\n📊 Testing data retrieval from Tutorias list...")
        tutorias_data = auth.get_list_items('Tutorias', ['ID', 'Aula', 'Estado'])
        if tutorias_data:
            print(f"✅ Retrieved {len(tutorias_data)} items from Tutorias list")
            if tutorias_data:
                print(f"  Sample item: {tutorias_data[0]}")
        else:
            print("❌ Failed to get data from Tutorias list")
            
    else:
        print("❌ Interactive authentication failed")
        
except ImportError as e:
    print(f"❌ Import error: {e}")
    print("Make sure you're running this from the project root directory")
except Exception as e:
    print(f"❌ Error: {e}")

print("\n🏁 Test completed")
