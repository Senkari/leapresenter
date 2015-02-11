using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;

using Microsoft.Win32;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

using Leap;

namespace PowerPointAddIn1
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class Overlay : Window, ILeapEventDelegate
    {
        //private:

        private Controller controller;
        private LeapEventListener listener;
        private PowerPoint.SlideShowWindow window;

        private bool isClosing = false;
        private bool slideShowActive;

        private BitmapImage penIcon;
        private BitmapImage highlighterIcon;
        private BitmapImage eraserIcon;

        private long timestamp                      = 0;
        private long deltaTime;
        private long slideSwitchCooldownTimer       = 0;         //slide can be changed repeatedly only after a specified waiting time
        private const long slideSwitchCooldownTime  = 1000000;   //in microseconds
        private long modeSwitchTimer                = 0;         //"thumb up - thumb down" -gesture triggers mode switching only if "thumb down" -part is performed quicly after "thumb down"
        private const long modeSwitchTime           = 1000000;   //in microseconds
        private bool thumbUpGestureActivated        = false;     //set to true when the first part, "thumb up", is triggered. 

        private enum Mode{
            Cursor, 
            Pen, 
            Highlighter,
            Eraser        
        };
        Mode mode           = Mode.Cursor;

        private void connectHandler()
        {
            this.controller.SetPolicy(Controller.PolicyFlag.POLICY_BACKGROUND_FRAMES);
            this.controller.EnableGesture(Gesture.GestureType.TYPE_SWIPE);
            //ENABLE ANY PREDEFINED GESTURES THAT ARE TO BE USED
        }

        private void newFrameHandler(Leap.Frame currentFrame)
        {
            if (slideShowActive)
            {

                deltaTime = currentFrame.Timestamp - timestamp;
                timestamp = currentFrame.Timestamp;
                slideSwitchCooldownTimer -= deltaTime;
                modeSwitchTimer -= deltaTime;
                if (slideSwitchCooldownTimer < 0) slideSwitchCooldownTimer = 0;
                if (modeSwitchTimer < 0) modeSwitchTimer = 0;

                if (!currentFrame.Hands.IsEmpty)
                {

                    FingerList indexList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_INDEX);
                    FingerList thumbList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_THUMB);
                    FingerList middleList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_MIDDLE);
                    FingerList ringList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_RING);
                    FingerList pinkyList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_PINKY);

                    // Get the first finger in the list of fingers
                    Finger index = indexList[0];
                    Finger thumb = thumbList[0];
                    Finger middle = middleList[0];
                    Finger ring = ringList[0];
                    Finger pinky = pinkyList[0];

                    updateCursor(index, ring, pinky);


                    if (thumbUpGestureActivated == false && thumb.IsExtended && !index.IsExtended && !middle.IsExtended && !ring.IsExtended && !pinky.IsExtended)
                    {
                        modeSwitchTimer = modeSwitchTime;
                        thumbUpGestureActivated = true;
                    }
                    if (thumbUpGestureActivated == true && modeSwitchTimer > 0 && !thumb.IsExtended && !index.IsExtended && !middle.IsExtended && !ring.IsExtended && !pinky.IsExtended)
                    {
                        thumbUpGestureActivated = false;

                        //set the next mode
                        if (mode == Mode.Cursor)        mode = Mode.Pen;
                        if (mode == Mode.Pen)           mode = Mode.Highlighter;
                        if (mode == Mode.Highlighter)   mode = Mode.Eraser;
                        if (mode == Mode.Eraser)        mode = Mode.Cursor;
                        updateIcons();
                    }

                    //ignore the gesture if left halfway while triggering time runs out
                    if (modeSwitchTimer == 0) thumbUpGestureActivated = false;
                 
                    //ignore the gesture if invalidated halfway
                    if (thumbUpGestureActivated == true && !(thumb.IsExtended && !index.IsExtended && !middle.IsExtended && !ring.IsExtended && !pinky.IsExtended)) thumbUpGestureActivated = false;

                    if (index.IsExtended && thumb.IsExtended)
                    {
                        if (mode == Mode.Pen) setPenMode();
                        if (mode == Mode.Highlighter) setHighlighterMode();
                        if (mode == Mode.Eraser) setEraserMode();
                    }
                    else setCursorMode();


                    GestureList gestures = currentFrame.Gestures();

                    for (int i = 0; i < gestures.Count(); i++)
                    {
                        Gesture gesture = gestures[i];

                        if (slideSwitchCooldownTimer == 0)
                        {

                            //gesture - action -mapping
                            if (horizontalSwipeToRight(gesture))
                            {
                                slideSwitchCooldownTimer = slideSwitchCooldownTime;
                                nextSlide();
                            }
                            else if (horizontalSwipeToLeft(gesture))
                            {
                                slideSwitchCooldownTimer = slideSwitchCooldownTime;
                                previousSlide();
                            }
                        }
                    }
                }
                else
                {
                    setCursorMode();
                }
            }
        }

        private void updateCursor(Finger index, Finger ring, Finger pinky)
        {           
            // Get the closest screen intercepting a ray projecting from the finger
            if(index.IsExtended && !ring.IsExtended && !pinky.IsExtended)
            {
                Screen screen = controller.LocatedScreens.ClosestScreenHit(index);

                if (screen != null && screen.IsValid)
                {
                    // Get the velocity of the finger tip
                    var tipVelocity = (int)index.TipVelocity.Magnitude;

                    // Use tipVelocity to reduce jitters when attempting to hold the cursor steady
                    if (tipVelocity > 0)
                    {
                        var xScreenIntersect = screen.Intersect(index, true).x;
                        var yScreenIntersect = screen.Intersect(index, true).y;

                        if (xScreenIntersect.ToString() != "NaN")
                        {
                            int x = (int)(xScreenIntersect * screen.WidthPixels);
                            int y = (int)(screen.HeightPixels - (yScreenIntersect * screen.HeightPixels));
                            MouseCursor.setCursor(x, y);                        
                        }
                    }
                }
            }
        }

        private void updateIcons()
        {
            pen.Margin          = new Thickness(0, 60, 0, 0);
            highlighter.Margin  = new Thickness(0, 165, 0, 0);
            eraser.Margin       = new Thickness(0, 270, 0, 0);           

            if (mode == Mode.Pen) pen.Margin                    = new Thickness(0, 60, 20, 0);
            if (mode == Mode.Highlighter) highlighter.Margin    = new Thickness(0, 165, 20, 0);
            if (mode == Mode.Eraser) eraser.Margin              = new Thickness(0, 270, 20, 0);
        }

        private bool horizontalSwipeToRight(Gesture gesture)
        {
            if (gesture.Type == Gesture.GestureType.TYPE_SWIPE)
            {
                SwipeGesture swipe = new SwipeGesture(gesture);
                if (swipe.Direction.x > 0.0f) return true;
            }
            return false;
        }

        private bool horizontalSwipeToLeft(Gesture gesture)
        {
            if (gesture.Type == Gesture.GestureType.TYPE_SWIPE)
            {
                SwipeGesture swipe = new SwipeGesture(gesture);
                if (swipe.Direction.x < 0.0f) return true;
            }
            return false;
        }

        private void nextSlide()
        {
            window.View.Next();
            canvas.Strokes.Clear();
        }

        private void previousSlide()
        {
            window.View.Previous();
            canvas.Strokes.Clear();
        }

        private void setPenMode()
        {
            var canvasSettings          = canvas.DefaultDrawingAttributes;    
            canvasSettings.StylusTip    = System.Windows.Ink.StylusTip.Ellipse;
            canvasSettings.Width        = 10;
            canvasSettings.Height       = 10;
            canvas.Cursor               = Cursors.Pen;
            canvas.EditingMode          = InkCanvasEditingMode.Ink;          
            MouseCursor.sendLeftMouseDown();         
        }

        private void setHighlighterMode()
        {
            var canvasSettings              = canvas.DefaultDrawingAttributes;
            canvasSettings.StylusTip        = System.Windows.Ink.StylusTip.Rectangle;
            canvasSettings.IsHighlighter    = true;
            canvasSettings.Color            = Colors.Yellow;
            canvasSettings.Width            = 10;
            canvasSettings.Height           = 20;
            canvas.Cursor                   = Cursors.Pen;
            canvas.EditingMode              = InkCanvasEditingMode.Ink;   
            MouseCursor.sendLeftMouseDown();
        }

        private void setEraserMode()
        {
            canvas.Cursor       = Cursors.Cross;
            canvas.EditingMode  = InkCanvasEditingMode.EraseByPoint;
            MouseCursor.sendLeftMouseDown();           
        }

        private void setCursorMode()
        {
            canvas.Cursor   = Cursors.Arrow;
            MouseCursor.sendLeftMouseUp();        
        }

        private void Overlay_Closing(object sender, EventArgs e)
        {
            this.isClosing = true;
            this.controller.RemoveListener(this.listener);
            this.controller.Dispose();
        }

        private BitmapImage LoadIcon(string name)
        {
            return new BitmapImage(new Uri(System.IO.Path.Combine(Environment.CurrentDirectory, "Leapresenter", "icons", name)));
        }
        //public:
   
        public Overlay()
        {
            InitializeComponent();
            this.controller= new Controller();
            this.listener = new LeapEventListener(this);
            controller.AddListener(listener);
            canvas.UseCustomCursor = true;
            canvas.Cursor = Cursors.Arrow;
            Closing += this.Overlay_Closing;

            penIcon = LoadIcon("pen_icon.png");
            highlighterIcon = LoadIcon("highlighter_icon.png");
            eraserIcon = LoadIcon("eraser_icon.png");
            
            pen.Source = penIcon;
            highlighter.Source = highlighterIcon;
            eraser.Source = eraserIcon;

            //Eraser
            canvas.EraserShape = new System.Windows.Ink.EllipseStylusShape(20, 20);

        }

        delegate void LeapEventDelegate(string EventName);

        public void LeapEventNotification(string EventName)
        {
            if (Dispatcher.CheckAccess())
            {
                switch (EventName)
                {
                    case "onInit":
                        break;
                    case "onConnect":
                        this.connectHandler();
                        break;
                    case "onFrame":
                        if (!this.isClosing)
                            this.newFrameHandler(this.controller.Frame());
                        break;
                }
            }
            else
            {
                Dispatcher.BeginInvoke(new LeapEventDelegate(LeapEventNotification), new object[] { EventName });
            }
        }

        public void setSlideShowWindow(PowerPoint.SlideShowWindow window)
        {
            this.window = window;
        }

        public void setSlideShowActive(bool slideShowActive)
        {
            this.slideShowActive = slideShowActive;
        }

    }

    public interface ILeapEventDelegate
    {
            void LeapEventNotification(string EventName);
    }

    public class LeapEventListener : Listener
    {
        ILeapEventDelegate eventDelegate;

        public LeapEventListener(ILeapEventDelegate delegateObject)
        {
            this.eventDelegate = delegateObject;
        }
        public override void OnInit(Controller controller)
        {
            this.eventDelegate.LeapEventNotification("onInit");
        }
        public override void OnConnect(Controller controller)
        {
            controller.SetPolicy(Controller.PolicyFlag.POLICY_BACKGROUND_FRAMES);
            controller.EnableGesture(Gesture.GestureType.TYPE_SWIPE);
            this.eventDelegate.LeapEventNotification("onConnect");
        }
        public override void OnFrame(Controller controller)
        {
            this.eventDelegate.LeapEventNotification("onFrame");
        }
        public override void OnExit(Controller controller)
        {
            this.eventDelegate.LeapEventNotification("onExit");
        }
        public override void OnDisconnect(Controller controller)
        {
            this.eventDelegate.LeapEventNotification("onDisconnect");
        }
    }
}
