cortina
=======

MediaMonkey Script: Tango DJ utility. Turn songs into cortinas without editing them.

More info: <https://www.tangoexchange.com/cortina-plugin-for-mediamonkey/>

**Installation:**

* [MediMonkey](https://www.mediamonkey.com/) must be installed on your system before proceeding.
* Download [cortina.mmip](https://github.com/wpkc/cortina/blob/master/cortina.mmip) and double-click the file, this will launch MediaMonkey and start the script installation process.
* Some versions of Windows will require entering an administration account name and password to proceed.
* Restart MediaMonkey.

**Updating:**

* From the MediaMonkey `Tools` menu, click the `Extensions` menu item.
* Click the `Find Updates` button at the bottom of the dialog box.
* If an update is available, click on `Cortina` in the list of extensions, then click the `Install Update` button.
* Some versions of Windows will require entering an administration account name and password to proceed.
* Restart MediaMonkey.

**Uninstall:**

* From the MediaMonkey `Tools` menu, click the `Extensions` menu item.
* Click on `Cortina` in the list of extensions.
* Click the `Uninstall` button.
* Click the `OK` button on the confirmation prompt.
* Select `Yes` or `No` option for the prompt asking if you wish to delete registry settings created by the script. Selecting `Yes` to delete registry settings will remove all traces of the cortina plugin; your cortina settings will no longer exist if you reinstall the cortina plugin later.
* Click `OK` on the uninstall information prompt.
* Restart MediaMonkey.

**Configuring:**

* From the `Play` menu, click the `Cortinas` menu item.
* Set the following options in the dialog box at any time (settings are re-read each time a cortina is found)
  * Select where to search for the word "cortina":
    - Choose where the script should search for the word `cortina` (not case-sensitive)
  * Cortina Length:
    - Drag slider to select the length of the entire cortina, including fade-in and fade-out time  (in seconds).
  * Fade In Time:
    - Drag slider to set a fade-in time (in seconds).
  * Fade Out Time:
    - Drag slider to set a fade-out time (in seconds).
  * Gap Time:
    - Drag slider to add an additional gap of silence after cortinas and songs (in seconds).
  * Cortina Volume:
    - Drag slider to set playback volume of cortinas (in percent of normal playback volume).
* Click the `Cancel` button to discard all changes to settings.
* Click the `Save` button to save new settings (new settings will be used on the next cortina found during playback).

**Build Notes:**

* To build the MediaMonkey Installation Package, compress all files, except for cortina.xml, with any Zip compressor utility. 
* Rename the resulting zip file to `cortina.mmip` (mmip = MediaMonkey Installer Package).
* Double-clicking on `cortina.mmip` will automatically start the MediaMonkey installation process.

**Additional notes:**

* To start playback of a cortina at a different point in the song:
    * Right-click on the cortina in the playlist or library pane.
    * Select `Properties` from the pop-up menu.
    * Click the `Details` tab in the song properties dialog box.
    * Change the `Start Time` setting to anything later that 0:00.
    * Click the `OK` button.
    * Note: setting the Fade-In option to 1 second or more will help hide any abrupt start to the cortina.
    
* If the song used for a cortina is shorter than the configured cortina length, the script will try to shorten the 
  full volume time period of the cortina.  If it is too short, it will be ignored and no cortina will be played.
  Take this into consideration if you use short silence files, they should not be labeled as cortinas when using this
  script or they might be skipped.

* For a uniform length of silence between each song follow this procedure:
  * From the `Tools` menu, click the `Options` menu item.
  * Click on `Output Plug-ins` in the list of options.
  * Click on the currently enabled output plug-in in the list of `Available Output Plug-ins:`
  * Checkmark the `Remove silence at the beginning / end of track` checkbox in your output plug-in.
  * Note that some output plug-ins, like `waveOut`, do not have this option. Use an output plug-in that does have this option, such as MediaMonkey DirectSound output.
  * Use the `Gap Time` slider in the Cortina options panel to add back in a fixed length of silence between all tracks.

* If you use `MediaMonkey WASAPI output` as your Output Plug-in and cortinas are not working:
    * Stop any playback in progress
    * Click Tools => Options
    * Click Player => Output Plugins
    * Under the `Available Output Plugins` select `MediaMonkey WASAPI output` then click the `Configure` button
    * Make sure the correct output device is selected from the list
    * Click the `Advanced` button.
    * Uncheck the `Event driven Exclusive mode (recommended)` checkbox
    * Uncheck the `Event driven Shared mode (recommended)` checkbox
    * Click the `OK` button on all open dialog boxes
    * Restart your playback
    * If this does not fix the problem, restore the original settings and consult your soundcard manufacturer for the correct settings.

