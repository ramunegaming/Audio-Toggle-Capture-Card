# Audio Toggle Capture Card

A small Python tool that fixes a recurring hardware quirk with a single hotkey — by talking directly to the Windows Core Audio API.

**Status:** Complete, in personal use

**Language:** Python (`comtypes`, `pycaw`)

## The problem

My capture card doesn't reliably start capturing audio when it's first plugged in. The manual fix is: open Sound settings → Properties → Listen tab → toggle "Listen to this device" off, then on again — every time, just to get audio working.

## How it works

Rather than automating UI clicks, this script manipulates the audio device's **property store directly via COM**, using the actual `PROPERTYKEY` GUID behind the "Listen to this device" feature (`{24DBB0FC-9311-4B3D-9CF0-18FF155639D4}`, undocumented by Microsoft). It:

- Opens the microphone endpoint's property store in read/write mode via `pycaw`/`comtypes`
- Reads which output device is currently set as the "listen" target
- Toggles the listen checkbox off/on directly at the property-store level (bypassing the Settings UI entirely)
- Cycles between two configured speaker devices, so the same hotkey can also switch the listen target between them

## Why it's here

This was one of my first coding projects — a genuinely practical solution to a mundane, everyday annoyance rather than a showcase piece. It's small, but it involved digging into an undocumented Windows COM interface to solve a problem the Settings UI couldn't fix reliably on its own. I still use it today.
