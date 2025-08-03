# BirdyFences
![birdyfences_sm](https://github.com/user-attachments/assets/ac577965-9f2f-483a-9bfd-4281b199ec73)
Thanks to Lycol50 for the icon

BirdyFences is a alternative to the StarDock's Fences.

![resim](https://github.com/user-attachments/assets/f6e8497d-b266-499a-b92b-2e62e5319b64)

# Install
* [Click here for all releases](https://github.com/HAKANKOKCU/BirdyFences/releases)
* You can also compile it yourself with .NET 9 SDK if you want.

# How to
Just install/open the executable to get started.

- To create a new fence, right click on the Title of a fence, then `New Fence`
- To remove a fence, right click in the Title of the fence, then `Remove Fence`
- To create a Portal Fence, right click in the Title of a fence then `New Portal Fence`, and select a folder to portal it.
- To Lock/Unlock a fence, right click on the Title of the fence, then `Lock Fence`, Which makes fence not moveable or editable.
- To edit title of the fence, double click on the title, then type the new title and press Enter, or ESC to undo.

# To automatically start when boot
## shell:startup
- Windows 11: Right Click into the exe file then `Show more options` then `Send to` and click `Desktop (Create Shortcut)`
- Windows 10: Right Click into the exe file then `Send to` and click `Desktop (Create Shortcut)`
- Do `CTRL+R` then type `shell:startup`
- Drag the `Birdy Fences - Shortcut` to the shell:startup file explorer
- And reboot to check if its working.
## Task scheduler
- Click on `Task scheduler library`(?) and click on `Create new task`(?). Name it what you want. DO NOT MAKE IT ADMIN.
- Click on `Triggers`(?) and `New`
- Select `When user logins`(?). Change others as you want
- Click on `Actions`(?) and `New`, Put app's path.
- Click on `Condintions`(?) disable `Only run on AC power`
- Click on `Settings`(?), disable the task timeout.
- Translations might be wrong.
