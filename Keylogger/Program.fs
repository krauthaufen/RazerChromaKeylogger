// Learn more about F# at http://fsharp.org

open System
open System.Windows.Forms
open System.Runtime.InteropServices
open System.Diagnostics
open Microsoft.FSharp.NativeInterop

module User32 =
    let SW_HIDE = 0
    let SW_RESTORE = 9
    let WH_KEYBOARD_LL = 13
    let WM_KEYDOWN = 0x0100n
    let WM_SYSKEYDOWN = 0x0104n
    let WM_KEYUP = 0x0101n
    let WM_SYSKEYUP = 0x0105n

    [<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
    extern IntPtr SetWindowsHookEx(int idHook, nativeint lpfn, nativeint hMod, uint32 dwThreadId);
    
    [<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
    extern bool UnhookWindowsHookEx(IntPtr hhk);

    
    [<DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
    extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam)

    [<DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)>]
    extern IntPtr GetModuleHandle(string lpModuleName)

    [<DllImport("kernel32.dll")>]
    extern IntPtr GetConsoleWindow()

    [<DllImport("user32.dll")>]
    extern bool ShowWindow(IntPtr hWnd, int nCmdShow)
    
    [<DllImport("user32.dll")>]
    extern int MapVirtualKeyEx(int code, int mapType, nativeint layout)
    
    [<DllImport("user32.dll")>]
    extern nativeint GetKeyboardLayout(int thread)

    type HookDelegate = delegate of int * nativeint * nativeint -> nativeint


    [<StructLayout(LayoutKind.Sequential)>]
    type KeyStruct =
        struct
            val mutable public vkCode : int
            val mutable public scanCode : int
            val mutable public flags : int
            val mutable public time : int
            val mutable public dwExtraInfo : nativeint
        end

    let setHook (callback : Keys -> unit) =
        let myHook = ref 0n
        let us = InputLanguage.FromCulture(Globalization.CultureInfo.CreateSpecificCulture "en-US")
        let dstLayout = us.Handle

        let down = System.Collections.Generic.HashSet<Keys>()

        let del = 
            HookDelegate (fun nCode wParam lParam ->
                if nCode >= 0 && (wParam = WM_KEYDOWN || wParam = WM_SYSKEYDOWN) then
                    let info = NativePtr.read (NativePtr.ofNativeInt<KeyStruct> lParam)
                    //let vkCode = Marshal.ReadInt32(lParam)
                    let scan = MapVirtualKeyEx(info.vkCode, 4, GetKeyboardLayout(0))
                    let vkCode = MapVirtualKeyEx(scan, 3, dstLayout)
                    let key = unbox<Keys> vkCode
                    if wParam = WM_SYSKEYDOWN && key = Keys.LControlKey then ()
                    elif down.Add key then callback key

                elif wParam = WM_KEYUP || wParam = WM_SYSKEYUP then
                    let info = NativePtr.read (NativePtr.ofNativeInt<KeyStruct> lParam)
                    let scan = MapVirtualKeyEx(info.vkCode, 4, GetKeyboardLayout(0))
                    let vkCode = MapVirtualKeyEx(scan, 3, dstLayout)
                    let key = unbox<Keys> vkCode
                    if wParam = WM_SYSKEYUP && key = Keys.LControlKey then ()
                    else down.Remove key |> ignore

                CallNextHookEx(!myHook, nCode, wParam, lParam)
            )
        let gc = GCHandle.Alloc(del)
        let ptr = Marshal.GetFunctionPointerForDelegate(del)
        use proc = Process.GetCurrentProcess()
        use curModule = proc.MainModule
        myHook := SetWindowsHookEx(WH_KEYBOARD_LL, ptr, GetModuleHandle(curModule.ModuleName), 0u)
        { new IDisposable with
            member x.Dispose() =
                if !myHook <> 0n then
                    UnhookWindowsHookEx(!myHook) |> ignore
                    gc.Free()
        }

open Colore
open Colore.Api
open Colore.Data
open Colore.Effects.Keyboard
open System.Globalization
open System.Threading
open Aardvark.Base
open System.Collections.Concurrent

module Translations =
    let toRazer =
        Dictionary.ofArray [|
            Keys.Back, Key.Backspace
            Keys.Tab, Key.Tab
            Keys.Enter, Key.Enter
            Keys.Pause, Key.Pause
            Keys.CapsLock, Key.CapsLock
            Keys.Escape, Key.Escape
            Keys.Space, Key.Space
            Keys.Prior, unbox<Key> 0 // TODO
            Keys.PageUp, Key.PageUp
            Keys.PageDown, Key.PageDown
            Keys.Next, unbox<Key> 0 // TODO
            Keys.End, Key.End
            Keys.Home, Key.Home
            Keys.Left, Key.Left
            Keys.Up, Key.Up
            Keys.Right, Key.Right
            Keys.Down, Key.Down
            Keys.Select, unbox<Key> 0 // TODO
            Keys.Print, unbox<Key> 0 // TODO
            Keys.Execute, unbox<Key> 0 // TODO
            Keys.Snapshot, unbox<Key> 0 // TODO
            Keys.PrintScreen, Key.PrintScreen
            Keys.Insert, Key.Insert
            Keys.Delete, Key.Delete
            Keys.Help, unbox<Key> 0 // TODO
            Keys.D0, Key.D0
            Keys.D1, Key.D1
            Keys.D2, Key.D2
            Keys.D3, Key.D3
            Keys.D4, Key.D4
            Keys.D5, Key.D5
            Keys.D6, Key.D6
            Keys.D7, Key.D7
            Keys.D8, Key.D8
            Keys.D9, Key.D9
            Keys.A, Key.A
            Keys.B, Key.B
            Keys.C, Key.C
            Keys.D, Key.D
            Keys.E, Key.E
            Keys.F, Key.F
            Keys.G, Key.G
            Keys.H, Key.H
            Keys.I, Key.I
            Keys.J, Key.J
            Keys.K, Key.K
            Keys.L, Key.L
            Keys.M, Key.M
            Keys.N, Key.N
            Keys.O, Key.O
            Keys.P, Key.P
            Keys.Q, Key.Q
            Keys.R, Key.R
            Keys.S, Key.S
            Keys.T, Key.T
            Keys.U, Key.U
            Keys.V, Key.V
            Keys.W, Key.W
            Keys.X, Key.X
            Keys.Y, Key.Y
            Keys.Z, Key.Z
            Keys.LWin, Key.LeftWindows
            Keys.RWin, unbox<Key> 0 // TODO
            Keys.Apps, unbox<Key> 0 // TODO
            Keys.Sleep, unbox<Key> 0 // TODO
            Keys.NumPad0, Key.Num0
            Keys.NumPad1, Key.Num1
            Keys.NumPad2, Key.Num2
            Keys.NumPad3, Key.Num3
            Keys.NumPad4, Key.Num4
            Keys.NumPad5, Key.Num5
            Keys.NumPad6, Key.Num6
            Keys.NumPad7, Key.Num7
            Keys.NumPad8, Key.Num8
            Keys.NumPad9, Key.Num9
            Keys.Multiply, Key.NumMultiply
            Keys.Add, Key.NumAdd
            Keys.Separator, Key.NumDecimal
            Keys.Subtract, Key.NumSubtract
            Keys.Decimal, Key.NumDecimal
            Keys.Divide, Key.NumDivide
            Keys.F1, Key.F1
            Keys.F2, Key.F2
            Keys.F3, Key.F3
            Keys.F4, Key.F4
            Keys.F5, Key.F5
            Keys.F6, Key.F6
            Keys.F7, Key.F7
            Keys.F8, Key.F8
            Keys.F9, Key.F9
            Keys.F10, Key.F10
            Keys.F11, Key.F11
            Keys.F12, Key.F12
            Keys.F13, unbox<Key> 0 // TODO
            Keys.F14, unbox<Key> 0 // TODO
            Keys.F15, unbox<Key> 0 // TODO
            Keys.F16, unbox<Key> 0 // TODO
            Keys.F17, unbox<Key> 0 // TODO
            Keys.F18, unbox<Key> 0 // TODO
            Keys.F19, unbox<Key> 0 // TODO
            Keys.F20, unbox<Key> 0 // TODO
            Keys.F21, unbox<Key> 0 // TODO
            Keys.F22, unbox<Key> 0 // TODO
            Keys.F23, unbox<Key> 0 // TODO
            Keys.F24, unbox<Key> 0 // TODO
            Keys.NumLock, Key.NumLock
            Keys.Scroll, Key.Scroll
            Keys.LShiftKey, Key.LeftShift
            Keys.RShiftKey, Key.RightShift
            Keys.LControlKey, Key.LeftControl
            Keys.RControlKey, Key.RightControl
            Keys.LMenu, Key.LeftAlt
            Keys.RMenu, Key.RightAlt
          
            Keys.OemSemicolon, Key.OemSemicolon
            Keys.Oem1, Key.OemSemicolon
            Keys.Oemplus, Key.OemEquals
            Keys.Oemcomma, Key.OemComma
            Keys.OemMinus, Key.OemMinus
            Keys.OemPeriod, Key.OemPeriod
            Keys.Oem2, unbox<Key> 0 // TODO
            Keys.OemQuestion, Key.OemSlash
            Keys.Oem3, unbox<Key> 0 // TODO
            Keys.Oemtilde, Key.OemTilde
            Keys.Oem4, unbox<Key> 0 // TODO
            Keys.OemOpenBrackets, Key.OemLeftBracket
            Keys.OemPipe, unbox<Key> 0 // TODO
            Keys.Oem5, Key.EurPound
            Keys.OemCloseBrackets, unbox<Key> 0 // TODO
            Keys.Oem6, Key.OemRightBracket
            Keys.OemQuotes, unbox<Key> 0 // TODO
            Keys.Oem7, Key.OemApostrophe
            Keys.Oem8, unbox<Key> 0 // TODO
            Keys.Oem102, unbox<Key> 0 // TODO
            Keys.OemBackslash, Key.EurBackslash
            Keys.ProcessKey, unbox<Key> 0 // TODO
            Keys.Packet, unbox<Key> 0 // TODO
            Keys.Attn, unbox<Key> 0 // TODO
            Keys.Crsel, unbox<Key> 0 // TODO
            Keys.Exsel, unbox<Key> 0 // TODO
            Keys.EraseEof, unbox<Key> 0 // TODO
            Keys.Play, unbox<Key> 0 // TODO
            Keys.Zoom, unbox<Key> 0 // TODO
            Keys.NoName, unbox<Key> 0 // TODO
            Keys.Pa1, unbox<Key> 0 // TODO
            Keys.OemClear, unbox<Key> 0 // TODO
        |]

    let tryGetRazerKey (k : Keys) =
        match toRazer.TryGetValue k with
        | (true, r) when int r <> 0 -> Some r
        | _ -> None

module ConcurrentDictionary =
    let toArray (d : ConcurrentDictionary<'a, 'b>) =
        let res = Array.zeroCreate d.Count
        use e = d.GetEnumerator()
        let mutable i = 0
        while e.MoveNext() && i < res.Length do 
            let kvp = e.Current
            res.[i] <- kvp.Key, kvp.Value
            i <- i + 1

        if i < res.Length then Array.take i res
        else res

type MyApp(init : unit -> unit) =
    inherit ApplicationContext()

    do init()

type MyItem(table : TableLayoutPanel) =
    inherit ToolStripControlHost(table)

    override x.GetPreferredSize(constrainTo) =
        constrainTo

type MyToolStripRenderer() =
    inherit ToolStripRenderer()

    override x.OnRenderItemText(e : ToolStripItemTextRenderEventArgs) =
        if e.Item.Selected then
            //let ne = ToolStripItemTextRenderEventArgs(e.Graphics, e.Item, e.Text, e.TextRectangle, Drawing.Color.Green, e.TextFont, Drawing.ContentAlignment.TopLeft)
            //base.OnRenderItemText(e)
            //let pt = Drawing.PointF(float32 e.TextRectangle.X, float32 e.TextRectangle.Y)
            //e.Graphics.DrawString(e.Text, e.TextFont, Drawing.Brushes.Black, pt)
            TextRenderer.DrawText(e.Graphics, e.Text, e.TextFont, e.TextRectangle, Drawing.Color.White, e.TextFormat)
        else
            base.OnRenderItemText(e)

    override x.OnRenderButtonBackground(e : ToolStripItemRenderEventArgs) =
        base.OnRenderButtonBackground(e)
        if e.Item.Selected then
            let color = Drawing.Color.FromArgb(255, 50, 50, 50)
            let rect = Drawing.Rectangle(0,0, e.Item.GetCurrentParent().Size.Width, e.Item.Size.Height)
            e.Graphics.FillRectangle(new Drawing.SolidBrush(color), rect)
            


[<EntryPoint; STAThread>]
let main argv =
    let info = AppInfo("Heatmapper", "KeyPressStatistics", "krauthaufen", "krauthaufen@awx.at", Category.Application)
    let api = Colore.ColoreProvider.CreateNativeAsync().Result
    api.InitializeAsync(info).Wait()

    let ico = 
        let name = typeof<MyApp>.Assembly.GetManifestResourceNames() |> Array.find (fun n -> n.EndsWith ".ico")
        use stream = typeof<MyApp>.Assembly.GetManifestResourceStream(name)
        new Drawing.Icon(stream)

    let stops =
        [|
            0.0, C3f.Blue
            0.3, C3f.Green
            0.6, C3f.Yellow
            1.0, C3f.Red
        |]

    let interpolate (t : float) (arr : array<float * 'a>) =
        let mutable l = 0
        let mutable r = arr.Length - 1

        while l <= r do
            let m = (l+r) / 2
            let (tm, vm) = arr.[m]
            if tm > t then
                r <- m - 1
            elif tm < t then
                l <- m + 1
            else
                l <- m
                r <- m - 1

        if r >= 0 && r < arr.Length then
            let (tl, vl) = arr.[r]
            let h = r + 1
            if h < arr.Length then
                let (tr, vr) = arr.[h]
                let tt = (t - tl) / (tr - tl)
                tt, vl, vr
            else
                0.0, vl, vl
        else
            let (_, vr) = arr.[l]
            0.0, vr, vr
            
            
    let min = 0.005
    let minExp = -log10 min

    let getColor (t : float) =
        let t = clamp 0.0 1.0 t

        let t = (minExp + log10 (max t min)) / minExp |> clamp 0.0 1.0

        let (t, a, b) = interpolate t stops
        let t = float32 t

        let a = a.ToHSLf()
        let b = b.ToHSLf()
        let f = HSLf(lerp a.H b.H t, lerp a.S b.S t, lerp a.L b.L t)
        let c = f.ToC3f()
        //let c = C3f(lerp a.R b.R t, lerp a.G b.G t, lerp a.B b.B t)
        //let c = hsl.ToC3f()

        //let c = HSVf(0.3f + float32 t * 0.7f, 1.0f, 1.0f).ToC3f()
        Color(c.R, c.G, c.B)

    let keyboard = api.Keyboard

    let mutable n = KeyboardCustom.Create()
    n.Set(getColor 0.0)

    let excluded = Set.ofList [Key.Space]
    let mutable maxCount = 0
    let mutable total = 0
    let mutable lastSave = 0
    let counts = ConcurrentDictionary<Key, int>()



    let file = Path.combine [Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData); "Keylogger"; "stats.csv"]
    let dir = System.IO.Path.GetDirectoryName file
    if not (System.IO.Directory.Exists dir) then System.IO.Directory.CreateDirectory dir |> ignore

    let razerKeys =
        let names = Enum.GetNames(typeof<Key>)
        let values = Enum.GetValues(typeof<Key>) |> unbox<Key[]>
        Dictionary.ofArray (Array.zip names values)

    let save (file : string) =  
        lastSave <- total
        let b = System.Text.StringBuilder()
        b.AppendLine(sprintf "# total: %d" total) |> ignore
        for (k, v) in counts |> ConcurrentDictionary.toArray |> Array.sortBy snd do
            b.AppendLine(sprintf "%A;%d" k v) |> ignore
        File.writeAllText file (b.ToString())

    let load (file : string) =
        if System.IO.File.Exists file then
            total <- 0
            maxCount <- 0
            counts.Clear()
            for l in File.readAllLines file do
                let p = l.Split(";")
                let mutable k = Key.A
                let mutable cnt = 0
                if p.Length = 2 && razerKeys.TryGetValue(p.[0], &k) && Int32.TryParse(p.[1], &cnt) then
                    counts.[k] <- cnt
                    if not (Set.contains k excluded) then
                        maxCount <- max maxCount cnt
                    total <- total + cnt

            lastSave <- total
        
    load file

    let updateColors() =
        let max = 
            if maxCount = 0 then 1
            else maxCount
        for KeyValue(k, c) in counts do
            n.[k] <- getColor (float c / float max)
        keyboard.SetCustomAsync(n).Wait()

    Thread.Sleep 1000
    updateColors()

    use __ = 
        User32.setHook (fun k ->
            match Translations.tryGetRazerKey k with
            | Some k ->
                let newCount = counts.AddOrUpdate(k, 1, fun _ o -> o + 1)
                if not (Set.contains k excluded) then
                    maxCount <- max maxCount newCount
                total <- total + 1
            | _ ->
                ()
            
            updateColors()

            if total - lastSave > 100 then
                save file
                
        )

    let icon = new NotifyIcon()
    let menu = new ContextMenuStrip()
    let close = new ToolStripButton("Close KeyLogger")
    //close.AutoSize <- false
    //close.Anchor <- AnchorStyles.Left ||| AnchorStyles.Right
    close.AutoToolTip <- false
    close.ToolTipText <- "Closes KeyLogger and saves the current stats"
    //close.TextAlign <- Drawing.ContentAlignment.MiddleLeft
    let s1 = new ToolStripLabel("total: 0")
    let original = s1.Font
    s1.Font <- new Drawing.Font(original.FontFamily, 14.0f, Drawing.FontStyle.Bold)

    let font = new Drawing.Font(original, Drawing.FontStyle.Bold)
    //let elems = 
    //    Array.init lines (fun _ ->
    //        let l = new ToolStripLabel("")
    //        l.Font <- font
    //        l
    //    )

    close.Click.Add (fun _ -> Application.Exit())
    
    menu.RenderMode <- ToolStripRenderMode.Professional
    menu.Renderer <- MyToolStripRenderer()
    menu.Items.Add(s1) |> ignore
    menu.Items.Add(new ToolStripSeparator()) |> ignore

    //table.Dock <- DockStyle.Fill

    let tablePos = menu.Items.Add(new ToolStripSeparator())

    menu.Items.Add(new ToolStripSeparator()) |> ignore



    let countString (c : int) =
        if c >= 1000000 then  sprintf "%.2fM" (float c / 1000000.0)
        elif c > 2000 then sprintf "%.1fk" (float c / 1000.0)
        else string c

    menu.Items.Add(close) |> ignore

    menu.Opening.Add (fun _ ->
        s1.Text <- sprintf "%s keystrokes" (countString total)
        let values = counts |> ConcurrentDictionary.toArray |> Array.sortByDescending snd
        let cnt = values.Length

        let table = new TableLayoutPanel()
        table.ColumnCount <- 2
        table.RowCount <- cnt
        table.MaximumSize <- Drawing.Size(1024, 300)
        //table.Dock <- DockStyle.Fill
        table.GrowStyle <- TableLayoutPanelGrowStyle.FixedSize
        table.AutoScroll <- true
        table.HorizontalScroll.Enabled <- false
        table.HorizontalScroll.Visible <- false
        table.Height <- table.GetRowHeights().[0] * 8
        for i in 0 .. cnt - 1 do
            if i < values.Length then
                let (k, c) = values.[i]

                let keyName = 
                    let keyName = string k
                    if keyName.StartsWith "Oem" then keyName.Substring 3
                    else keyName

                let header = new Label(Text = keyName)
                let value = new Label(Text = countString c)
                
                header.TextAlign <- Drawing.ContentAlignment.MiddleRight
                value.TextAlign <- Drawing.ContentAlignment.MiddleLeft
                value.Font <- font
                table.Controls.Add header
                table.Controls.Add value
                table.SetCellPosition(header, TableLayoutPanelCellPosition(0, i))
                table.SetCellPosition(value, TableLayoutPanelCellPosition(1, i))
              
        
        table.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 10.0f)) |> ignore
        table.ColumnStyles.Add(ColumnStyle(SizeType.AutoSize, 80.0f)) |> ignore
        menu.Items.RemoveAt(tablePos)
        menu.Items.Insert(tablePos, new ToolStripControlHost(table))

    )

    menu.BackColor <- Drawing.Color.FromArgb(255, 30, 30, 30)
    menu.ForeColor <- Drawing.Color.FromArgb(255, 220, 220, 220)
    menu.ShowImageMargin <- false
    
    
    icon.Icon <- ico
    icon.ContextMenuStrip <- menu
    icon.Visible <- true
    

    Application.ApplicationExit.Add (fun _ -> 
        icon.Visible <- false
        api.UninitializeAsync().Wait()
    )
    Application.Run()
    save file

    0
