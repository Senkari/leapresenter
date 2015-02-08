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

        private int timestamp;
        private int deltaTime;
        private int cooldownTimeLeft = 0;
        private const int cooldownTime = 1000000;   //in milliseconds

        enum Mode{
            Cursor, 
            Pen, 
            Eraser
        };
        Mode mode;  //not needed currently

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

                deltaTime = (int)currentFrame.Timestamp - timestamp;
                timestamp = (int)currentFrame.Timestamp;
                cooldownTimeLeft -= deltaTime;
                if (cooldownTimeLeft < 0) cooldownTimeLeft = 0;

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

                    GestureList gestures = currentFrame.Gestures();

                    for (int i = 0; i < gestures.Count(); i++)
                    {
                        Gesture gesture = gestures[i];

                        if (cooldownTimeLeft == 0)
                        {

                            //gesture - action -mapping
                            if (horizontalSwipeToRight(gesture))
                            {
                                cooldownTimeLeft = cooldownTime;
                                nextSlide();
                            }
                            else if (horizontalSwipeToLeft(gesture))
                            {
                                cooldownTimeLeft = cooldownTime;
                                previousSlide();
                            }
                            else if (thumb.IsExtended)
                            {
                                setPenMode(canvas);
                            }
                            else if (middle.IsExtended)
                            {
                                setEraserMode(canvas);
                            }
                            else if ((thumb.IsExtended && middle.IsExtended) || (!thumb.IsExtended && !middle.IsExtended))
                            {
                                setCursorMode(canvas);
                            }
                        }
                    }
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
                    if (tipVelocity > 25)
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

        private void setPenMode(InkCanvas canvas)
        {
            canvas.Cursor = Cursors.Pen;
            canvas.EditingMode = InkCanvasEditingMode.Ink;
            MouseCursor.sendLeftMouseDown();
            mode = Mode.Pen;
        }

        private void setEraserMode(InkCanvas canvas)
        {
            canvas.Cursor = Cursors.Cross;
            canvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
            MouseCursor.sendLeftMouseDown();
            mode = Mode.Eraser;
        }

        private void setCursorMode(InkCanvas canvas)
        {
            canvas.Cursor = Cursors.Arrow;
            MouseCursor.sendLeftMouseUp();
            mode = Mode.Cursor;
        }

        private void Overlay_Closing(object sender, EventArgs e)
        {
            this.isClosing = true;
            this.controller.RemoveListener(this.listener);
            this.controller.Dispose();
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

            //Canvas settings
            var canvasSettings = canvas.DefaultDrawingAttributes;

            //Pen     
            canvasSettings.StylusTip = System.Windows.Ink.StylusTip.Ellipse;
            canvasSettings.Width = 10;
            canvasSettings.Height = 10;

            //Eraser
            canvas.EraserShape = new System.Windows.Ink.EllipseStylusShape(15, 15);
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
