# Troubleshooting Outlook Add-in Installation

## Common "Installation Failed" Issues & Solutions

### Issue 1: Manifest Validation Errors
**Problem:** The manifest XML has formatting or validation issues
**Solution:** Use the simplified manifest

### Issue 2: HTTPS Requirements  
**Problem:** Outlook requires HTTPS URLs, can't load local files
**Solution:** Host the file online first

### Issue 3: Browser/Platform Specific Issues
**Problem:** Different Outlook versions have different requirements

## 🚀 **EASIEST SOLUTION: Try These Working Alternatives**

### Option A: Use GitHub Pages (Recommended)
1. **Push your code to GitHub**
2. **Enable GitHub Pages** in repository settings
3. **Update manifest URL** to: `https://yourusername.github.io/authopsy/dist/taskpane.html`
4. **Upload updated manifest**

### Option B: Use CodePen (Instant Testing)
1. **Go to [codepen.io](https://codepen.io)**
2. **Create new pen**
3. **Copy the HTML content** from `dist/taskpane.html`
4. **Get the debug URL** (looks like: `https://codepen.io/pen/debug/abcdef`)
5. **Use the manifest-simple.xml** and replace the URL

### Option C: Use a Free Hosting Service
- **Netlify**: Drag & drop the `dist` folder
- **Vercel**: Connect your GitHub repo  
- **GitHub Pages**: Enable in repo settings
- **Surge.sh**: Run `npx surge dist/`

## 🔧 **Manual Testing Steps**

### For Outlook Web:
1. Go to **Settings** ⚙️ > **View all Outlook settings**
2. **General** > **Manage add-ins**
3. **Add a custom add-in** > **Add from file**
4. Select `manifest-simple.xml` (the new one I created)
5. **Accept any security warnings**

### For Outlook Desktop:
1. **File** > **Manage Add-ins** (or **Insert** > **Get Add-ins**)
2. **My add-ins** > **Add a custom add-in** > **Add from file**
3. Browse to `manifest-simple.xml`
4. **Install**

## 🐛 **If Still Failing, Try This:**

### Check Manifest Validation:
1. Go to [Office Add-in Validator](https://dev.office.com/add-in-validator)
2. Upload your manifest file
3. Fix any reported errors

### Enable Developer Mode:
1. **Outlook Web**: Press `F12` to open developer tools
2. Look for **console errors** when installing
3. Check **Network tab** for failed requests

### Try Minimal Test:
Create a super simple test manifest that just shows "Hello World"

## 📝 **Quick Fix Manifest**

I've created `manifest-simple.xml` with these improvements:
- ✅ **Proper HTTPS URLs** (placeholder - you'll update these)
- ✅ **Simplified structure**
- ✅ **Standard icons from web**
- ✅ **Shorter text** (some versions have character limits)
- ✅ **Valid XML formatting**

## 🎯 **Next Steps:**

1. **Try the `manifest-simple.xml`** first
2. **If that works**, we know the format is correct
3. **Host your HTML file online** (GitHub Pages is easiest)
4. **Update the URLs** in the manifest
5. **Reinstall** the add-in

## 🆘 **Emergency Backup Plan:**

If nothing works, I can create a version that works as:
- **Outlook Web Extension** (different approach)
- **Browser Bookmarklet** (works in any email interface)
- **Desktop App** (Electron-based)

**What would you prefer to try first?**
1. Host on GitHub Pages?
2. Try the simplified manifest?
3. Use CodePen for instant testing?
4. Try a different approach entirely?
