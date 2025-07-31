# How to Update the Application Icon

## Method 1: Using Project File Configuration (Recommended)

### Step 1: Create Your Icon File
1. **Create or obtain your icon image** (PNG, JPG, etc.)
2. **Convert to .ico format** using one of these methods:
   - **Online converters**: ConvertICO, IcoConvert, etc.
   - **Image editing software**: Photoshop, GIMP, Paint.NET
   - **Command line tools**: ImageMagick

### Step 2: Prepare the Icon File
Your .ico file should include multiple sizes:
- **16x16 pixels** (for taskbar)
- **32x32 pixels** (for desktop shortcuts)
- **48x48 pixels** (for file explorer)
- **256x256 pixels** (for high DPI displays)

### Step 3: Add Icon to Project
1. **Place your icon file** in the `ADAUDIT` folder
2. **Name it** `app.ico` (or update the reference below)
3. **Update the project file**:

```xml
<PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net6.0-windows</TargetFramework>
    <UseWindowsForms>true</UseWindowsForms>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <ApplicationIcon>app.ico</ApplicationIcon>  <!-- Add this line -->
    
                    <AssemblyTitle>SonicADRecon</AssemblyTitle>
            <AssemblyDescription>LDAP Domain Management Tool</AssemblyDescription>
            <AssemblyCompany>ADAUDIT</AssemblyCompany>
            <AssemblyProduct>SonicADRecon</AssemblyProduct>
    <AssemblyVersion>1.0.0.0</AssemblyVersion>
    <FileVersion>1.0.0.0</FileVersion>
</PropertyGroup>
```

### Step 4: Update Form Icon (Optional)
Add this line to `MainForm.Designer.cs` in the `InitializeComponent()` method:

```csharp
this.Icon = new System.Drawing.Icon("app.ico");
```

## Method 2: Using Embedded Resources

### Step 1: Add Icon as Resource
1. **Right-click** on your project in Solution Explorer
2. **Add** → **Existing Item**
3. **Select your .ico file**
4. **Set Build Action** to "Resource"

### Step 2: Update Project File
```xml
<ItemGroup>
    <Resource Include="app.ico" />
</ItemGroup>
```

### Step 3: Load Icon in Code
```csharp
// In MainForm.Designer.cs
this.Icon = new System.Drawing.Icon(GetType().Assembly.GetManifestResourceStream("ADAUDIT.app.ico"));
```

## Method 3: Using System Icons

### Option A: Use Windows System Icons
```csharp
// In MainForm.Designer.cs
this.Icon = System.Drawing.SystemIcons.Application;
// or
this.Icon = System.Drawing.SystemIcons.Information;
// or
this.Icon = System.Drawing.SystemIcons.Shield;
```

### Option B: Use Custom Icon from File
```csharp
// In MainForm.Designer.cs
if (File.Exists("app.ico"))
{
    this.Icon = new System.Drawing.Icon("app.ico");
}
```

## Icon Creation Tools

### Free Online Tools:
- **ConvertICO.com** - Convert images to .ico
- **IcoConvert.com** - Online icon converter
- **Favicon.io** - Create icons from text/images

### Desktop Software:
- **GIMP** (Free) - Image editing with icon export
- **Paint.NET** (Free) - Simple image editing
- **Photoshop** (Paid) - Professional image editing
- **IconWorkshop** (Paid) - Professional icon creation

### Command Line:
```bash
# Using ImageMagick
magick convert image.png -resize 16x16 icon16.ico
magick convert image.png -resize 32x32 icon32.ico
magick convert image.png -resize 48x48 icon48.ico
```

## Icon Design Guidelines

### Best Practices:
1. **Keep it simple** - Icons should be recognizable at small sizes
2. **Use consistent style** - Match your application's design theme
3. **Test at different sizes** - Ensure it looks good at 16x16, 32x32, etc.
4. **Use transparent backgrounds** - Icons look better with transparency
5. **Follow Windows guidelines** - Use standard icon conventions

### Recommended Colors:
- **Primary**: #0078D4 (Windows Blue)
- **Secondary**: #107C10 (Windows Green)
- **Accent**: #D83B01 (Windows Orange)
- **Background**: Transparent or white

### Icon Themes for SonicADRecon:
- **Shield with checkmark** - Security/Audit theme
- **Database icon** - LDAP/Active Directory theme
- **User with key** - User management theme
- **Network nodes** - Domain/Network theme

## Troubleshooting

### Common Issues:
1. **"Icon stream is not in the expected format"**
   - Ensure your .ico file is valid
   - Try recreating the icon file
   - Check file corruption

2. **Icon not showing in taskbar**
   - Rebuild the application
   - Clear Windows icon cache
   - Restart the application

3. **Icon not showing in file explorer**
   - Ensure .ico file is in project directory
   - Check ApplicationIcon property in .csproj
   - Rebuild and redeploy

### Verification Steps:
1. **Build the project** - `dotnet build`
2. **Check executable** - Look for icon in file explorer
3. **Run application** - Verify icon in taskbar
4. **Check properties** - Right-click exe → Properties → Icon

## Example Icon Implementation

### Complete Setup:
1. **Create icon file**: `app.ico` in project root
2. **Update .csproj**:
```xml
<ApplicationIcon>app.ico</ApplicationIcon>
```
3. **Update MainForm.Designer.cs**:
```csharp
this.Icon = new System.Drawing.Icon("app.ico");
```
4. **Build and test**:
```bash
dotnet build
dotnet run
```

This will give your SonicADRecon application a professional, branded icon that appears in the taskbar, file explorer, and application window. 