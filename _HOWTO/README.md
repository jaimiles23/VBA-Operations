# HOWTO 
This repository focuses on Macros for your Personal Macrobook. Personal Macros are used in your day-to-day operations and are applicable to a wider scope of projects. Contrast this with Macros written within a specific file that can only be accessed when that file is open. I focus on Personal Macros given their breadth and general usefulness for business users.

## Personal Macrobook
VBA Macros are tied to workbooks. For instance, a user may write a VBA macro within an Excel workbook and then that macro is saved within that Workbook.  Office-365 Applications also contain a **Personal Workbook** A personal macrobook is tied to your Office Account and is accessible between files of the same application. A macro in your Personal Excel Macrobook can be used in two separate files without the other being open. This HOWTO focuses on Excel, as it is the application with the most macro development. The same principles are applicable in other standard Office Suite Products, i.e., Word & PPT. 

### Creating a Personal Workbook
You can easily create a Personal workbook by recoridng a macro and choosing to save the macro in your "Personal Workbook"
![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/howto/StoreInPersonalWb.png?raw=true)

Your Excel personal workbook is available below - you will almost never need to access it. Office-365 applications will automatically load the workbook with the application:
    `C:\Users\User Name\AppData\Roaming\Microsoft\Excel\XLSTART`


## Customization 

### Renaming Modules
It's just as important to modulate the code in your Personal Workbook as it is any other project. In fact, I'll argue it's more important as VBA Business Users run myriad macros on highly variable projects, with month+ timeframes between runs.

To rename a module, 
1. Select the Module
2. Go to View -> Properties Window
3. Rename the Module
   1. Only use alphanumeric characters.


![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/howto/RenameModule.png?raw=true)


### GUI Environment
The Office VBA IDE is a harsh environment to develop it. I've been tempted on numerous occassions to switch back & forth with VS Code. However, there are undeniable benefits to use the native IDE. For instance, direct access to the data. There are a number of ways to make the enviornment more suitable.

#### Bigger Font
My eyes hurt looking at the VBA IDE:

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/howto/IDE_UnadjustedFont.png?raw=true)

There's no intuitive `cntrl + +` or `cntrl + -` shortcuts to adjust zoom of your VBA IDE. In fact, there is no zoom function at all so far as I can tell! You can adjust the size of the text displayed. 
> Tools -> Options -> Editor Format

I recommend sticking to the majority of default options until there's a specific need not to. 
![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/howto/Options_FontEditor.png?raw=true)

I've found size 16 works well for my aging eyes. 
![.](https://raw.githubusercontent.com/jaimiles23/VBA-Operations/main/_images/howto/IDE_LargerFont.png?raw=true)

Font colors for keywords, variables, etc. are established conventions within the VBA community. Changing may confuse collaborators.

#### Debugging Console
With any coding project, it's also imperative to add a debugging console. VBA refers to this as the **Immediate Window**. Toggle this with either the shortcut `cntrl + g` or:
> View -> Immediate Window

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/howto/HelloWorld.png?raw=true)

