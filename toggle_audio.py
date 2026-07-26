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
microphone_endpoint_id = "{0.0.1.00000000}.{e819d0a6-0cd8-4962-93b1-c8f19ecc6c30}"

speaker_devices = [
    {
        "name": "Speakers (Creative Pebble Pro)",
        "endpoint_id": "{0.0.0.00000000}.{c4f83cbc-fb21-4ede-a494-94a01656c651}"
    },
    {
        "name": "Speakers (USB Audio Device)",
        "endpoint_id": "{0.0.0.00000000}.{f5e1f4b4-6ce2-4492-978a-123ae4b06b2f}"
    }
]


def get_device_store():
    enumerator = AudioUtilities.GetDeviceEnumerator()
    dev = enumerator.GetDevice(microphone_endpoint_id)
    return dev.OpenPropertyStore(STGM_READWRITE)


def get_current_listening_device(store):
    device_pk = PROPERTYKEY()
    device_pk.fmtid = GUID(LISTEN_SETTING_GUID)
    device_pk.pid = LISTENING_DEVICE_PID
    try:
        value = store.GetValue(device_pk)
        if value and value.GetValue():
            current_id = value.GetValue()
            for speaker in speaker_devices:
                if current_id == speaker["endpoint_id"]:
                    return speaker
    except:
        pass
    return speaker_devices[0]


def get_next_speaker(current_speaker):
    if current_speaker == speaker_devices[0]:
        return speaker_devices[1]
    return speaker_devices[0]


def set_listening_checkbox(property_store, value: bool):
    checkbox_pk = PROPERTYKEY()
    checkbox_pk.fmtid = GUID(LISTEN_SETTING_GUID)
    checkbox_pk.pid = CHECKBOX_PID

    new_value = PROPVARIANT(VT_BOOL)
    new_value.union.boolVal = value
    property_store.SetValue(checkbox_pk, new_value)


def set_listening_device(property_store, output_device):
    device_pk = PROPERTYKEY()
    device_pk.fmtid = GUID(LISTEN_SETTING_GUID)
    device_pk.pid = LISTENING_DEVICE_PID

    new_value = PROPVARIANT(VT_LPWSTR)
    new_value.union.pwszVal = output_device["endpoint_id"]
    property_store.SetValue(device_pk, new_value)


def toggle_device(store, times=1):
    for _ in range(times):
        set_listening_checkbox(store, False)
        set_listening_checkbox(store, True)


def main():
    try:
        store = get_device_store()
        current_speaker = get_current_listening_device(store)
        next_speaker = get_next_speaker(current_speaker)
        toggle_device(store, times=2)
        set_listening_device(store, next_speaker)
        set_listening_checkbox(store, True)
    except Exception:
        pass


if __name__ == "__main__":
    main()