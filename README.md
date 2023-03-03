# <center>My Case Tracker



[toc]

![image-20230223101621913](./README.assets/image-20230223101621913.png)



## The difference between My CaseTracker  and My Case Tracker Rating

> Select script based on your need

- **My Case Tracker** 

  It will generate CSV with **5** columns:

  | Date     | Time     | Case          | Type    | Note                |
  | -------- | -------- | ------------- | ------- | ------------------- |
  | 1/1/2023 | 12:12:00 | Incident88888 | Written | Win10 Upgrade issue |

  

- **My Case Tracker Rating**

  It will generate CSV with **4** columns:

  | Date     | Time     | Case          | Rating  |
  | -------- | -------- | ------------- | ------- |
  | 1/1/2023 | 12:12:00 | Incident88888 | 10 |



## How to Use:

Right click->Run with Powershell



## How to Change Daily Target:

> Only available in **My Case Tracker** 

Modify`$target` to under **variables** block, default value:  **35** 

```powershell
$target = 35
```



## How to Switch CSV on Local/OneDrive path:

> Do not change this option after CSV is created, or it may create a new CSV.

Change `$KeepLocal` to `$true` or `$false` under **variables** block, default value: `$false`

```powershell
[Bool]$KeepLocal = $false
```



## Where is CSV Stored?

By default, the CSV is stored on **Documents** folder

>  If OneDrive is installed, it will set with OneDrive-Documents folder

You can also use the 'open CSV folder via file explorer' option to open the folder contains this CSV



## What if I Want Change The Name of CSV?

Change `$Filename`line to  under **variables** block

```bash
$Filename = <custom_csv_name>
```



## What if I Want Change the CSV to Custom Path?

1.  Set`$KeepLocal` to $true under **variables** block

   ```bash
   [Bool]$KeepLocal = $true
   ```

   

2. Find function **Get-ProfilePath**, modify the `$profilepath` in main **else** block

```bash
    else{
        # Modify custom path here with $keeplocal set to $true.
        $profilepath = <your_custom_path>
    }
```



![image-20230222165259511](./README.assets/image-20230222165259511.png)
