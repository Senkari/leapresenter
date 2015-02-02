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

        private Boolean isClosing = false;
        private Boolean slideShowActive;
        private Boolean penMode = false;
        private Boolean eraserMode = false;


        private PowerPoint.SlideShowWindow window;

        private int prevX;
        private int prevY;
        private long currentTime;
        private long previousTimeMouse;
        private long previousTimeGesture;
        private long deltaTimeMouse;
        private long deltaTimeGesture;

        private void connectHandler()
        {
            this.controller.SetPolicy(Controller.PolicyFlag.POLICY_BACKGROUND_FRAMES);
            this.controller.EnableGesture(Gesture.GestureType.TYPE_SWIPE);
        }

        private void newFrameHandler(Leap.Frame currentFrame)
        {
            if (slideShowActive)
            {
                currentTime = currentFrame.Timestamp;
                deltaTimeMouse = currentTime - previousTimeMouse;

                

                if (deltaTimeMouse > 10000)
                {

                    previousTimeMouse = currentTime;

                    if (!currentFrame.Hands.IsEmpty)
                    {
                        FingerList indexList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_INDEX);
                        FingerList thumbList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_THUMB);
                        FingerList middleList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_MIDDLE);
                        FingerList ringList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_RING);
                        FingerList pinkyList = currentFrame.Fingers.FingerType(Finger.FingerType.TYPE_PINKY);

                        // Get the first finger in the list of fingers
                        Finger finger = indexList[0];
                        Finger finger1 = thumbList[0];
                        Finger middle = middleList[0];
                        Finger ring = ringList[0];
                        Finger pinky = pinkyList[0];

                        // Get the closest screen intercepting a ray projecting from the finger
                        if(finger.IsExtended && !ring.IsExtended && !pinky.IsExtended){

                            Screen screen = controller.LocatedScreens.ClosestScreenHit(finger);

                            if (screen != null && screen.IsValid)
                            {
                                // Get the velocity of the finger tip
                                var tipVelocity = (int)finger.TipVelocity.Magnitude;

                                // Use tipVelocity to reduce jitters when attempting to hold
                                // the cursor steady
                                if (tipVelocity > 25)
                                {
                                    var xScreenIntersect = screen.Intersect(finger, true).x;
                                    var yScreenIntersect = screen.Intersect(finger, true).y;

                                    if (xScreenIntersect.ToString() != "NaN")
                                    {
                                        int x = (int)(xScreenIntersect * screen.WidthPixels);
                                        int y = (int)(screen.HeightPixels - (yScreenIntersect * screen.HeightPixels));

                                        MouseCursor.setCursor(x, y);

                                        if (finger1.IsExtended)
                                        {
                                            if (!penMode)
                                            {
                                                var canvasSettings = canvas.DefaultDrawingAttributes;
                                                canvasSettings.StylusTip = System.Windows.Ink.StylusTip.Ellipse;
                                                canvasSettings.Width = 10;
                                                canvasSettings.Height = 10;
                                                canvas.EditingMode = InkCanvasEditingMode.Ink;
                                                MouseCursor.sendLeftMouseDown();
                                                penMode = true;
                                            }

                                        }
                                        else if (middle.IsExtended)
                                        {
                                            if (!eraserMode)
                                            {
                                                canvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
                                                canvas.EraserShape = new System.Windows.Ink.EllipseStylusShape(15, 15);
                                                if (penMode)
                                                {
                                                    penMode = false;
                                                }
                                                else MouseCursor.sendLeftMouseDown();

                                                eraserMode = true;
                                            }
                                            
                                        }
                                        else if (penMode || eraserMode)
                                        {
                                            MouseCursor.sendLeftMouseUp();
                                            penMode = false;
                                            eraserMode = false;
                                        }
                                        prevX = x;
                                        prevY = y;
                                    }
                                }
                            }
                        }
                    }
                }

                GestureList gestures = currentFrame.Gestures();

                for (int i = 0; i < gestures.Count(); i++)
                {
                    Gesture gesture = gestures[i];
                    deltaTimeGesture = currentTime - previousTimeGesture;


                    if (gesture.Type == Gesture.GestureType.TYPE_SWIPE && deltaTimeGesture > 1000000)
                    {
                        SwipeGesture swipe = new SwipeGesture(gesture);

                        previousTimeGesture = currentTime;

                        if (swipe.Direction.x > 0.0f)
                        {
                            window.View.Next();
                            //canvas.Strokes.Clear();
                            return;
                        }
                        else if (swipe.Direction.x < 0.0f)
                        {
                            window.View.Previous();
                            //canvas.Strokes.Clear();
                            return;
                        }

                    }
                }
            }
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
            canvas.Cursor = Cursors.Pen;
            Closing += this.Overlay_Closing;
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
