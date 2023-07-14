
from pyuac import main_requires_admin
from tkinter import *
import socket
import random
import string
import subprocess
import re
import json
import netifaces
import wmi
from tkinter import ttk
import sv_ttk

@main_requires_admin
def main():

    #HostName
    def random_windows_hostname():
        random_string = ''.join(random.choices(string.ascii_uppercase + string.digits, k=7))
        return f"DESKTOP-{random_string}"

    def change_hostname(new_hostname):
        if validate_hostname_entry(None, new_hostname):
            # Run the command to change the hostname
            command = f'wmic computersystem where name="%COMPUTERNAME%" call rename name="{new_hostname}"'
            output = subprocess.check_output(command, shell=True)
    #-----------------------------------------------------------------------------------

    #MAC Address
    # the registry path of network interfaces
    network_interface_reg_path = r"HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e972-e325-11ce-bfc1-08002be10318}"
    # the transport name regular expression, looks like {AF1B45DB-B5D4-46D0-B4EA-3E18FA49BF5F}
    transport_name_regex = re.compile("{.+}")
    # the MAC address regular expression


    def get_random_mac_address():
        """Generate and return a MAC address in the format of WINDOWS"""
        # Get the hexdigits 'A' to 'F' and numbers '0' to '9'
        hexdigits_and_numbers = 'ABCDEF0123456789'
        # 2nd character must be 2, 4, A, or E
        second_character = random.choice("24AE")
        # Generate the remaining 10 characters
        remaining_characters = ''.join(random.choices(hexdigits_and_numbers, k=10))
        # Return the MAC address
        return ''.join(random.choices(hexdigits_and_numbers)) + second_character + remaining_characters

    def get_adapter_name_name_from_mac(mac_address):
        c = wmi.WMI()
        # Query for network adapters
        adapters = c.Win32_NetworkAdapter(PhysicalAdapter=True)
        for adapter in adapters:
            if adapter.MACAddress.lower() == mac_address.lower():
                return adapter.Name

        # Return None if the MAC address is not found
        return None

    def clean_mac(mac):
        mac_cleaned = re.sub(r'[^a-fA-F0-9]', '', mac)
        return mac_cleaned 

    def get_interface_info():
        interfaces = netifaces.interfaces()
        interface_list = []

        for interface in interfaces:
            addresses = netifaces.ifaddresses(interface)

            if netifaces.AF_LINK in addresses and addresses[netifaces.AF_LINK][0]['addr'] != '':
                mac_address = addresses[netifaces.AF_LINK][0]['addr']

                # Format MAC address to match the format in the output
                mac_address_formatted = re.sub(r':', '-', mac_address).upper()

                # Check if MAC address exists in "getmac" command output
                output = subprocess.check_output('getmac', universal_newlines=True)
                if mac_address_formatted in output:
                    interface_list.append((mac_address, interface))

        return interface_list

    def disable_adapter(adapter_index):
        # use wmic command to disable our adapter so the MAC address change is reflected
        disable_output = subprocess.check_output(f"wmic path win32_networkadapter where index={adapter_index} call disable").decode()
        return disable_output

    def enable_adapter(adapter_index):
        # use wmic command to enable our adapter so the MAC address change is reflected
        enable_output = subprocess.check_output(f"wmic path win32_networkadapter where index={adapter_index} call enable").decode()
        return enable_output

    def get_original_mac_address(current_mac):
        output = subprocess.check_output(f"reg QUERY " +  network_interface_reg_path.replace("\\\\", "\\")).decode()
        for interface in re.findall(rf"{network_interface_reg_path}\\\d+", output):
            interface_content = subprocess.check_output(f"reg QUERY {interface.strip()}").decode()
            pattern = r"OriginalNetworkAddress\s+REG_SZ\s+(.+)"
            match = re.search(pattern, interface_content)
            if match:
                original_mac = match.group(1)
                if clean_mac(current_mac).lower() == clean_mac(re.search(r"NetworkAddress\s+REG_SZ\s+(\S+)", interface_content).group(1)).lower():
                    return original_mac
        return None

    def convert_number(number):
        if isinstance(number, float) and not number.is_integer():
            return number
        else:
            return int(number)

    def convert_speed(speed_bps):
        speed_mbps = int(speed_bps)/ 1000000  # Convert to Mbps
        speed_gbps = int(speed_bps) / 1000000000  # Convert to Gbps
        if speed_mbps > 1000:
            speed_mbps = 0
        if speed_gbps > 1000:
            speed_gbps = 0
        return (speed_gbps, "Gbps") if speed_gbps >=1 else (speed_mbps, "Mbps")

    def get_full_info(mac_address):
        c = wmi.WMI()
        # Query for network adapters
        adapters = c.Win32_NetworkAdapter(PhysicalAdapter=True)
        for adapter in adapters:
            if adapter.MACAddress.lower() == mac_address.lower():
                return adapter
        # Return None if the MAC address is not found
        return None

    def get_adapters_list():
        adapters = []
        for adapter in get_interface_info():
            LS = []
            LS.append(get_full_info(adapter[0]).ProductName)
            LS.append(adapter[0].upper())
            LS.append(get_full_info(adapter[0]).NetEnabled)
            changed = True if get_original_mac_address(clean_mac(adapter[0])) != clean_mac(adapter[0]) and get_original_mac_address(clean_mac(adapter[0])) is not None else False
            LS.append(changed)
            LS.append(str(convert_number(float(convert_speed(get_full_info(adapter[0]).Speed)[0]))) + " " + convert_speed(get_full_info(adapter[0]).Speed)[1])
            adapters.append(LS)
        return adapters

    def on_focus_in(event):
        if entry_new_mac.get() == "XX-XX-XX-XX-XX-XX":
            entry_new_mac.delete(0, END)

    def change_mac_address(adapter_transport_name, new_mac_address):
        # use reg QUERY command to get available adapters from the registry
        output = subprocess.check_output(f"reg QUERY " +  network_interface_reg_path.replace("\\\\", "\\")).decode()
        for interface in re.findall(rf"{network_interface_reg_path}\\\d+", output):
            # get the adapter index
            adapter_index = int(interface.split("\\")[-1])
            interface_content = subprocess.check_output(f"reg QUERY {interface.strip()}").decode()
            if adapter_transport_name in interface_content:
                if checkbox_var.get() == 1:
                    disable_adapter(adapter_index)
                    # if the transport name of the adapter is found on the output of the reg QUERY command
                    # then this is the adapter we're looking for
                    # change the MAC address using reg ADD command
                    enable_adapter(adapter_index)
                changing_mac_output = subprocess.check_output(f"reg add {interface} /v NetworkAddress /d {clean_mac(new_mac_address)} /f").decode()
                break

    #------------------------------------------------------------------------------------------- 
    window = Tk()
    window.title("AutoMac - by Darkonex")
    window.geometry("550x550")
    sv_ttk.set_theme("dark")
    window.iconbitmap("./icon.ico")
    #--------------------------------------------------------
    #HostName
    #get original hostname from json or get new one
    try:
        with open("./settings.json", 'r') as json_file:
            data = json.load(json_file)
            original_hostname = data.get('original Hostname')
            checkbox_value = data.get('checkbox')
    except FileNotFoundError:
        original_hostname = socket.gethostname()
        checkbox_value = 0
        with open("./settings.json", 'w') as json_file:
            json.dump({'original Hostname': original_hostname, 'checkbox': 0}, json_file)
    def update_checkbox_value(new_value):
        with open("./settings.json", 'r+') as json_file:
            data = json.load(json_file)
            data['checkbox'] = new_value
            json_file.seek(0)  # Move the file pointer to the beginning
            json.dump(data, json_file)
            json_file.truncate()
    #print the original hostname to entry_newhostname
    def print_original_hostname():
        entry_newhostname.delete(0, END)
        entry_newhostname.insert(0, original_hostname)
        validate_hostname_entry(None, original_hostname)
    #print the Random hostname to entry_newhostname
    def print_random_hostname_to_entry():
        random_hostname = random_windows_hostname()
        entry_newhostname.delete(0, END)
        entry_newhostname.insert(0, random_hostname)
        validate_hostname_entry(None, random_hostname)
    #"entry_newhostname" validation
    def validate_hostname_entry(event, new_value):
        if not new_value.isupper():
            new_value = new_value.upper()
            entry_newhostname.delete(0, END)
            entry_newhostname.insert(0, new_value)
        if event:
            if event.state == 4 or event.state == 12:  # Ctrl key has a state value of 4
                return
        pattern = r'^DESKTOP-[A-Z0-9]{7}$'  # Update the pattern to allow for up to 7 characters
        if re.match(pattern, new_value):
            entry_newhostname.config(background="white")
            button_change_hostname.config(state='active')
            return True
        else:
            entry_newhostname.config(background="red")
            button_change_hostname.config(state='disabled')
            return False
    #Hostname frame
    frame = ttk.LabelFrame(window, text="HostName")
    frame.grid(column=0, row=0, padx=10, pady=10, sticky="w")
    # Create a label for "Hostname"
    label_hostname = ttk.Label(frame, text="Hostname:")
    label_hostname.grid(column=0, row=0 ,padx=5, pady=5, sticky="w")
    # Create a label to display the actual hostname
    label_hostname_value = ttk.Label(frame, text=socket.gethostname())
    label_hostname_value.grid(column=1, row=0, sticky="w")
    #label_hostname_value.pack(side=LEFT, anchor="nw", padx=5, pady=5)
    # Create a label for "new Hostname"
    label_hostname = ttk.Label(frame, text="New Hostname:")
    label_hostname.grid(column=0, row=1 ,padx=5, pady=5, sticky="w")
    #create an Entry for "New Hostname"
    entry_newhostname = ttk.Entry(frame)
    entry_newhostname.bind("<KeyRelease>", lambda event: validate_hostname_entry(event, entry_newhostname.get()))
    entry_newhostname.bind("<FocusOut>", lambda event: validate_hostname_entry(event, entry_newhostname.get()))
    entry_newhostname.grid( column=1, row=1, padx=5)
    #create restore original hostname button
    button_restore_original_hostname = ttk.Button(frame, text="Restore Original", command= print_original_hostname)
    button_restore_original_hostname.grid(column=0, row=2, padx=5, pady=5, sticky="w")
    #create "random" button
    button_random = ttk.Button(frame, text="Generate Random", command=print_random_hostname_to_entry)
    button_random.grid(column=1, row=2, padx=5, pady=5, sticky="w", ipadx=9)
    #create change hostname button
    #"frame"
    ff = ttk.Frame(frame)
    button_change_hostname = ttk.Button(ff, text="Change", command= lambda: change_hostname(entry_newhostname.get()))
    button_change_hostname.grid(column=0, row=0, sticky="w", padx=5, pady=5)
    ff.grid(column=2, row=1, ipadx=66, padx=3)
    #--------------------------------------------------------
    #MAC
    def convert_to_format(string):
        sliced_parts = [string[i:i+2] for i in range(0, len(string), 2)]
        formatted_string = "-".join(sliced_parts)
        return formatted_string
    def print_random_mac():
        entry_new_mac.delete(0, END)
        entry_new_mac.insert(0, convert_to_format(get_random_mac_address()))
        validate_MAC_entry(None, entry_new_mac.get())
    def validate_MAC_entry(event, new_value):
        if not new_value.isupper():
            new_value = new_value.upper()
            entry_new_mac.delete(0, END)
            entry_new_mac.insert(0, new_value)
        if event:
            if event.state == 4 or event.state == 12:  # Ctrl key has a state value of 4
                return
        pattern = r'^([0-9A-F]{2}-){5}[0-9A-F]{2}$'  # Pattern for MAC address validation
        if re.match(pattern, new_value):
            entry_new_mac.config(background="white")
            change_mac_button.config(state="active")
            #Enable change button
            return True
        else:
            entry_new_mac.config(style='style.TEntry', background="red")
            change_mac_button.config(state="disabled")
            #Disable change button
            return False
    def select(event):
        global originalMAC
        global currentMAC
        try:
            restore_original_button.config(state="normal")
            originalMAC = get_original_mac_address(clean_mac(table.item(table.focus())['values'][1]))
            if originalMAC == None:
                originalMAC = table.item(table.focus())['values'][1]
            currentMAC = table.item(table.focus())['values'][1]
            label_selected_adapter.config(text="selected adapter: " + table.item(table.focus())['values'][0])
        except Exception as err: pass
    
    tbf = ttk.Frame(window)
    table = ttk.Treeview(tbf, height=3)
    scrollbar = ttk.Scrollbar(tbf, orient=VERTICAL, command=table.yview)
    table.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(column=1, row=0, sticky="ns")
    table["columns"] = ("name", "MAC", "status", "changed", "speed")
    table.column("#0", width=0, stretch=NO)
    table.column("name", width=250, stretch=NO)
    table.column("MAC", width=100, anchor="center", stretch=NO)
    table.column("status", width=45, anchor="center", stretch=NO)
    table.column("changed", width=60, anchor="center", stretch=NO)
    table.column("speed", width=60, anchor="center", stretch=NO)
    table.heading("#0", text="")
    table.heading("name", text="Adapter Name")
    table.heading("MAC", text="MAC Adress", anchor="center")
    table.heading("status", text="Status", anchor="center")
    table.heading("changed", text="Changed", anchor="center")
    table.heading("speed", text="Speed", anchor="center")
    table.bind("<<TreeviewSelect>>", select)  
    table.grid(column=0, row=0, sticky="w", pady=5)
    def update_treeview():
        try: table.delete(*table.get_children())
        except: pass
        for adapter in get_adapters_list():
            table.insert("", "end", values=adapter)
    update_treeview()
    table.bind("<<TreeviewSelect>>", select)
    tbf.grid(column=0, row=1)

    #main frame
    MAC_frame = ttk.LabelFrame(window, text="MAC", height=500)

    #Selected adapter Label
    label_selected_adapter = ttk.Label(MAC_frame, text="Selected Adapter:")
    label_selected_adapter.grid(column=0, row=0, padx=5, pady=5, sticky='w')
    
    #newmacframe
    newmacframe = ttk.Frame(MAC_frame)
    newmacframe.grid(column=0, row=1, padx=5, pady=5, sticky='w')
    #New Mac label
    label_new_mac = ttk.Label(newmacframe, text="New MAC:")
    label_new_mac.grid(column=0, row=0, padx=5, pady=5, sticky='w')
    #New MAC entry
    entry_new_mac = ttk.Entry(newmacframe)
    entry_new_mac.delete(0, END)
    entry_new_mac.insert(0, "XX-XX-XX-XX-XX-XX")
    entry_new_mac.bind("<FocusIn>", on_focus_in)
    entry_new_mac.bind("<KeyRelease>", lambda event: validate_MAC_entry(event, entry_new_mac.get()))
    entry_new_mac.bind("<FocusOut>", lambda event: validate_MAC_entry(event, entry_new_mac.get()))
    entry_new_mac.grid(column=1, row=0, padx=5, pady=5, sticky='w')
    #Generate Random MAC Button
    random_MAC_button = ttk.Button(newmacframe, text="Genereate MAC", command=print_random_mac)
    random_MAC_button.grid(column=2, row=0, padx=5, pady=5, sticky='w')
    #Restore original and generate MAC buttonb frame
    RandG_frame = ttk.Frame(MAC_frame)
    RandG_frame.grid(column=0, row=2, padx=5, pady=5, sticky='w', ipadx=105)
    #Restore original button
    def restore_original_mac():
        entry_new_mac.delete(0, END)
        if originalMAC:
            entry_new_mac.insert(0, originalMAC.upper().replace(":", "-"))
        else:
            entry_new_mac.insert(0, originalMAC.upper().replace(":", "-"))
    restore_original_button = ttk.Button(RandG_frame, text="Restore Original", command= restore_original_mac ,state="disabled", width=15)
    restore_original_button.pack(side="left", padx=78, pady=5)
    #check box
    checkbox_var = IntVar()
    checkbox = ttk.Checkbutton(MAC_frame, text="Automatically restart the adapter(s) to apply changes", variable=checkbox_var, command=lambda: update_checkbox_value(checkbox_var.get()))
    #make it retrive check value from settings
    checkbox_var.set(checkbox_value)
    checkbox.grid(column=0, row=3, padx=5, pady=5, sticky='w')
    def changeMAC():
        transportname = ""
        for interface in get_interface_info():
            if clean_mac(currentMAC).lower() == clean_mac(interface[0]).lower():
                transportname = interface[1]
                break
        change_mac_address(transportname, entry_new_mac.get())
        update_treeview()
    def changeall(event):
        for interface in get_interface_info():
            change_mac_address(interface[1], get_random_mac_address())
        if event:
            update_treeview()
    #change and cahnge all button frame
    change_button_frame = ttk.Frame(MAC_frame)
    change_button_frame.grid(column=0, row=4, padx=5, pady=5, sticky='w')
    #change button
    change_mac_button = ttk.Button(change_button_frame, text="Change", command=changeMAC)
    change_mac_button.grid(column=0, row=4, padx=5, pady=5, sticky='w')
    #change all button
    change_all_mac_button = ttk.Button(change_button_frame, text="Change All", command=changeall)
    change_all_mac_button.grid(column=1, row=4, padx=5, pady=5, sticky='w')
    MAC_frame.grid(column=0, row=2, padx=5, pady=5, sticky='w')
    
    window.mainloop()
    
if __name__ == "__main__":
    main()