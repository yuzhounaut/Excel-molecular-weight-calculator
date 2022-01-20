This repository provides code for the Molecular Weight V4.1 Excel add-in. After installing, users can get the molecular weight of a string by calling the new Excel function mw(chemicalName), e.g. typing in a cell =mw("H2O") would yield 18.01528. This add-in can handle isotopes, charges, and complex formulas. This add-in also allows users to specify custom names for both chemicals (e.g. "water" for H<sub>2</sub>O) and specific functional groups (e.g. "Me" for CH<sub>3</sub>).

Recommended reading:

https://fiehnlab.ucdavis.edu/projects/seven-golden-rules/accurate-mass

https://www.livescience.com/20581-weigh-atom.html

https://iupac.qmul.ac.uk/AtWt/

How to use:
1. Download the add-in file "MWAddInV4.xlam".
2. Put the add-in file in the C:\Users\\"username"\AppData\Roaming\Microsoft\AddIns folder.
3. Enable the add-in. Click the **File** tab, click **Options**, and then click the **Add-Ins** category. In the **Manage** box, click **Excel Add-ins**, and then click **Go**. The **Add-Ins** dialog box appears. In the **Add-Ins available** box, select the check box next to the add-in that you want to activate, and then click **OK**.
4. Example of monoisotopic molecular mass calculation, both formulas can be used.
5. An example of calculating the molecular weight of an isotope, taking <sup>13</sup>C as an example, need to add a "!" symbol in front.

![11](https://user-images.githubusercontent.com/86154919/150274880-58c52dd8-7caf-4280-8079-cf3b076bcbad.png)



Excel精确分子量计算插件
如何使用：
1. 下载"MWAddInV4.xlam"文件。
2. 将"MWAddInV4.xlam"文件放在C:\Users\\"username"\AppData\Roaming\Microsoft\AddIns文件夹，"username"处填写你的用户名。
3. 激活Excel加载项。单击“文件”选项卡，单击“选项”，然后单击“加载项”类别。在“管理”框中，单击“Excel 加载项”，然后单击“转到”。将显示 "加载项" 对话框。在"可用加载项"框中，选中要激活的加载项MWAddInV4旁边的复选框，然后单击"确定"。
4. 单一同位素分子质量计算示例，两公式均可使用。
5. 计算同位素分子量示例，以<sup>13</sup>C为例，需在前面加“!”符号。

![11](https://user-images.githubusercontent.com/86154919/150274881-71572c1c-def0-4d77-9c30-ed3aebfca660.png)





The codes were originally created at http://www.sciencechatforum.com/viewtopic.php?f=15&t=17657, but the site is now inaccessible.

The following description comes from http://www.sciencechatforum.com/viewtopic.php?f=15&t=26204&start=0 created by Natural ChemE, and the site is also inaccessible now.

Short version for power users:
1. Open Excel to a new workbook.
2. Ensure that the "Developer" tab is showing in the ribbon. If it's not, enable it (see instructions in older thread or Google how to enable the Developer tab in Excel).
3. Go to the Visual Basic button on the Developer tab.
4. Right-click the new Excel workbook in the Project treeview, then make a new module.
5. Optionally, rename the new module "MWAddInV4". (Doesn't really affect anything.)
6. Download text file with code, then copy/paste it into the module.
7. Save the workbook as an Excel Add-In with a name like "MWAddInV4.xlam".
8. Close the workbook.
9. Open a new Excel window (existing file or new one, doesn't matter.)
10. Enable the add-in. (See old thread or Google how to enable an Excel add-in.)
11. Now add-in will work for all Excel spreadsheets on your computer until you disable or/and delete the add-in.
12. If you intend to share a spreadsheet using these functions, either have the recipient install this add-in or put the add-in code in a module in the file to be shared and save that file as a macro-enabled workbook (*.xlsm).

Notes:
1. **Basic use**: In Excel, type =*mw("H2O")* for the molecular weight of water.
2. **Names go in quotes**: =*mw("H2")* gets the molecular weight of molecular hydrogen, but =*mw(H2)* gets the molecular weight of whatever is in cell H2 on the spreadsheet.
3. **Syntax errors return -999999**: If -999999 is returned, then the code detected an error. Try checking the input to ensure that there are no errors.
4. **New "MostCommonIsotope" function**: =*MostCommonIsotope("O")* returns 16 since the most common isotope of Oxygen has a mass number of 16.
5. **Case sensitive**: =*mw("CO")* gets for carbon monoxide (CO), but =*mw("Co")* gets for cobalt (Co).
6. **Complex formulas work**: =*mw("[P(C4H9)3C14H29]Cl")* does correctly return the molecular weight for [P(C4H9)3C14H29]Cl. In general you can have as parentheses, occurrences of elements, isotopes, charges, etc. as you want. The code doesn't care if you use ()'s or []'s; it seems them as the same. While not recommended as it's poor form, you can even mix and match, e.g. =*mw("(H]2")* does work.
7. **Subscripts can be fractions or decimals**: Subscripts can be fractions or decimals, e.g. =*mw("H0.5")* or =*mw("H1/2")* will correctly read as H<sub>0.5</sub> or H<sub>1/2</sub> and return half the molecular weight of Hydrogen.
8. **Charges affect mass**: Hydroxide, OH-, is calculated as the molecular weight of Oxygen plus Hydrogen plus an electron. This is slightly more than the mass of OH alone. Positive charges slightly reduce mass.
9. **Charges are their own atomic elements**: Charge signs (- and +) are treated as their own atomic elements. For a charge of negative two, type in "-2" instead of "2-".
10. **You can specify specific isotopes**: Put an exclamation point followed by the atomic number of an isotope before the atomic symbol. Example: *mw("!2H2O")* is heavy water, D<sub>2</sub>O. *mw("D2O")* would also work.
11. **Prevalence-weighted average is used by default when no isotope is specified**: *mw("H2O")* doesn't specify any isotopes, so both Hydrogen and Oxygen use prevalence-weighted averages. (If you don't know what this means, it's okay because it's probably want you want.)
12. **Changing default isotope behavior**: Don't like prevalence weighting? Use *mw("H2O", True)* instead. The second argument can be True or False, and if not supplied is False by default. True uses the most common isotope whenever isotope isn't specified while False uses the prevalence-weighted average whenever isotope isn't specified. True or False doesn't matter when for symbols which already have an isotope specified.
Put an exclamation point followed by the atomic number of an isotope before the atomic symbol. Example:
13. **Charges affect mass**: Hydroxide, OH<sup>-</sup>, is calculated as the molecular weight of Oxygen plus Hydrogen plus an electron. This is slightly more than the mass of OH alone. Positive charges slightly reduce mass.

Advanced notes:
1. **You can add in custom names**: See comments in VBA code in function MWICustom. Custom names for overall chemicals go into MWICustom. Custom names for charges, isotopes, functional groups, and other elements of chemical formulas go into AMUCustom.
2. **Updates**: See this thread for future updates. VBA code contains commented-out C# code which can generate new VBA tables from NIST's data. C# code can be run in Microsoft's Visual Studio or, probably, Mono for Linux users.
3. **Bugs**: If you find any bugs, please report them to this thread!
4. **Feature requests**: If you would like any new features, please report them to this thread!

Historical notes:
1. **There used to be an "Isotope" version**: The add-in defines two functions MW() and MWI() that both do the same thing. This is because there used to be two different add-ins; MW() was small and did only prevalence-weighted stuff while MWI() was larger but did isotopes too. Technically the new version is the "Isotope" version and the smaller version of the code has been discontinued.
2. **Versions before 4.0 often returned (slightly) incorrect results**: I made a mistake interpreting NIST's notation. For example, I thought that the current atomic weight of Hydrogen, 1.00794(7), meant "1.007947 where the last '7' isn't significant". However this was incorrect; it actually meant "1.00794 ± 0.00007". Version 4.0 corrects this misinterpretation, returning correct values.
3. **Not all older versions included changes in mass due to charge**: This add-in used to disregard charge, e.g. OH and OH<sup>-</sup> both had the same molecular weight. Versions 4.0 and later consistently consider charges as increasing molecular mass by , e.g. OH<sup>-</sup> outweighs OH by one electron mass. This also means that you can get negative results in extreme cases, e.g. =mw("+") will return the negative of an electron's mass.
4. **!2D and !3T no longer work**: Hydrogen's second isotope can be explicitly referred to using either "!2H" or "D", but not "!2D". Hydrogen's third isotope can be explicitly referred to using either "!3H" or "T", but not "!3T". Some older versions accepted "!2D" and "!3T" because D and T were regarded as elements, but now they're regarded as custom isotope names instead; it doesn't make sense to specify the isotope of an isotope.
