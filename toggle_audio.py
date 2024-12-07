from typing import Optional
from comtypes import GUID
from comtypes.automation import VT_BOOL, VT_LPWSTR, VT_EMPTY
from comtypes.persist import STGM_READWRITE
from pycaw.api.mmdeviceapi import PROPERTYKEY
from pycaw.api.mmdeviceapi.depend import PROPVARIANT
from pycaw.utils import AudioUtilities

# Hardcoded values
LISTEN_SETTING_GUID = "{24DBB0FC-9311-4B3D-9CF0-18FF155639D4}"
CHECKBOX_PID = 1
LISTENING_DEVICE_PID = 0

# Device configuration
microphone_name = 'CU4K30'
speaker_devices = ['Speakers (Creative Pebble Pro)', 'Speakers (USB Audio Device)']

def get_current_listening_device(store):
    device_pk = PROPERTYKEY()
    device_pk.fmtid = GUID(LISTEN_SETTING_GUID)
    device_pk.pid = LISTENING_DEVICE_PID
    
    try:
        value = store.GetValue(device_pk)
        if value and value.GetValue():
            current_guid = value.GetValue()
            # Find the device name from GUID
            output_devices = get_list_of_active_coreaudio_devices("output")
            for device in output_devices:
                if device.id == current_guid:
                    return device.FriendlyName
    except:
        pass
    return speaker_devices[0]  # Default to first speaker if not found

def get_next_speaker(current_speaker):
    if current_speaker == speaker_devices[0] or current_speaker not in speaker_devices:
        return speaker_devices[1]
    return speaker_devices[0]

def main():
    try:
        # Get current state
        store = get_device_store(microphone_name)
        if store is None:
            print("Failed to open property store")
            exit(1)

        # Get current speaker and determine next speaker
        current_speaker = get_current_listening_device(store)
        next_speaker = get_next_speaker(current_speaker)
        
        # Toggle sequence
        set_listening_checkbox(store, False)  # First disable listening
        set_listening_device(store, next_speaker)  # Switch to next speaker
        set_listening_checkbox(store, True)  # Enable listening
        
    except ValueError as e:
        print(f"Error: {str(e)}")
        exit(1)
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        exit(1)

def set_listening_checkbox(property_store, value: bool):
    checkbox_pk = PROPERTYKEY()
    checkbox_pk.fmtid = GUID(LISTEN_SETTING_GUID)
    checkbox_pk.pid = CHECKBOX_PID

    new_value = PROPVARIANT(VT_BOOL)
    new_value.union.boolVal = value
    property_store.SetValue(checkbox_pk, new_value)

def set_listening_device(property_store, output_device_name: Optional[str]):
    if output_device_name is not None:
        listening_device_guid = get_GUID_from_name(output_device_name)
    else:
        listening_device_guid = None

    device_pk = PROPERTYKEY()
    device_pk.fmtid = GUID(LISTEN_SETTING_GUID)
    device_pk.pid = LISTENING_DEVICE_PID

    if listening_device_guid is not None:
        new_value = PROPVARIANT(VT_LPWSTR)
        new_value.union.pwszVal = listening_device_guid
    else:
        new_value = PROPVARIANT(VT_EMPTY)

    property_store.SetValue(device_pk, new_value)

def get_device_store(device_name: str):
    device_guid = get_GUID_from_name(device_name)
    enumerator = AudioUtilities.GetDeviceEnumerator()
    dev = enumerator.GetDevice(device_guid)
    store = dev.OpenPropertyStore(STGM_READWRITE)
    return store

def get_GUID_from_name(device_name: str) -> str:
    input_devices = get_list_of_active_coreaudio_devices("input")
    for device in input_devices:
        if device_name.lower() in device.FriendlyName.lower():
            return device.id
    output_devices = get_list_of_active_coreaudio_devices("output")
    for device in output_devices:
        if device_name.lower() in device.FriendlyName.lower():
            return device.id
    raise ValueError(f"Device '{device_name}' not found!")

def get_list_of_active_coreaudio_devices(device_type: str) -> list:
    import comtypes
    from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator, EDataFlow, DEVICE_STATE
    from pycaw.constants import CLSID_MMDeviceEnumerator

    if device_type != "output" and device_type != "input":
        raise ValueError("Invalid audio device type.")

    if device_type == "output":
        EDataFlowValue = EDataFlow.eRender.value
    else:
        EDataFlowValue = EDataFlow.eCapture.value

    devices = list()
    device_enumerator = comtypes.CoCreateInstance(
        CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER)
    if device_enumerator is None:
        raise ValueError("Couldn't find any devices.")
    collection = device_enumerator.EnumAudioEndpoints(EDataFlowValue, DEVICE_STATE.ACTIVE.value)
    if collection is None:
        raise ValueError("Couldn't find any devices.")

    count = collection.GetCount()
    for i in range(count):
        dev = collection.Item(i)
        if dev is not None:
            if not ": None" in str(AudioUtilities.CreateDevice(dev)):
                devices.append(AudioUtilities.CreateDevice(dev))

    return devices

if __name__ == "__main__":
    main()
