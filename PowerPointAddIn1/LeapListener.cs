/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

using Leap;

namespace PowerPointAddIn1
{
    class LeapListener : Listener
    {
        //private:
        private PowerPoint.SlideShowWindow window;
        private Overlay overlayWindow;
        private System.Windows.Shapes.Polyline line;
        private bool slideShowActive;
        private bool penMode = false;
        private int prevX;
        private int prevY;
        private long currentTime;
        private long previousTimeMouse;
        private long previousTimeGesture;
        private long deltaTimeMouse;
        private long deltaTimeGesture;
    
        //public:
        public override void OnConnect(Controller controller)
        {
            //System.Console.WriteLine("LeapMotion connected");
            controller.EnableGesture(Gesture.GestureType.TYPE_SWIPE);
        }

        public override void OnDisconnect(Controller controller)
        {
            //System.Console.WriteLine("LeapMotion disconnected");
        }
        public override void OnFrame(Controller controller)
        {

            if (slideShowActive)
            {

                Frame currentFrame = controller.Frame();

                currentTime = currentFrame.Timestamp;
                deltaTimeMouse = currentTime - previousTimeMouse;

                if (deltaTimeMouse > 10000)
                {

                    previousTimeMouse = currentTime;

                    if (!currentFrame.Hands.IsEmpty)
                    {

                        // Get the first finger in the list of fingers
                        Finger finger = currentFrame.Fingers[0];

                        // Get the closest screen intercepting a ray projecting from the finger
                        Screen screen = controller.LocatedScreens.ClosestScreenHit(finger);

                        if (finger.Type() == Finger.FingerType.TYPE_INDEX || finger.Type() == Finger.FingerType.TYPE_MIDDLE)
                        {
                            if (screen != null && screen.IsValid)
                            {
                                // Get the velocity of the finger tip
                                var tipVelocity = (int)finger.TipVelocity.Magnitude;

                                // Use tipVelocity to reduce jitters when attempting to hold
                                // the cursor steady
                                if (tipVelocity > 15)
                                {
                                    var xScreenIntersect = screen.Intersect(finger, true).x;
                                    var yScreenIntersect = screen.Intersect(finger, true).y;

                                    if (xScreenIntersect.ToString() != "NaN")
                                    {
                                        int x = (int)(xScreenIntersect * screen.WidthPixels);
                                        int y = (int)(screen.HeightPixels - (yScreenIntersect * screen.HeightPixels));

                                        MouseCursor.setCursor(x, y);
                                        if (penMode && (prevX != x || prevY != y))
                                        {
                                            //Action draw(() => overlayWindow.draw(x, y));
                                            overlayWindow.Dispatcher.Invoke(new Action(() => overlayWindow.addPointToLine(line, x, y)));
                                        }
                                        prevX = x;
                                        prevY = y;
                                    }
                                }
                            }
                        }
                    }
                }

                GestureList gestures = controller.Frame().Gestures();

                for (int i = 0; i < gestures.Count(); i++)
                {
                    Gesture gesture = gestures[i];
                    deltaTimeGesture = currentTime - previousTimeGesture;

                    _anothercommentstartshere
                    if (gesture.Type == Gesture.GestureType.TYPE_SWIPE && deltaTimeGesture > 500000)
                    {
                        SwipeGesture swipe = new SwipeGesture(gesture);

                        previousTimeGesture = currentTime;

                        if (swipe.Direction.x > 0.0f)
                        {
                            window.View.Next();
                        }
                        else if (swipe.Direction.x < 0.0f)
                        {
                            window.View.Previous();
                        }

                    }
                    _anothercommentendshere

                    if (gesture.Type == Gesture.GestureType.TYPE_SWIPE && deltaTimeGesture > 500000)
                    {
                        previousTimeGesture = currentTime;
                        if (!penMode)
                        {
                            line = new System.Windows.Shapes.Polyline();
                            overlayWindow.Dispatcher.Invoke(new Action(() => overlayWindow.addPolyline(line)));
                            //window.View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
                            penMode = true;
                        }
                        else
                        {
                            //MouseSimulator.ReleaseLeftMouseButton();
                            //window.View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerArrow;
                            penMode = false;
                        }
                        
                    }
                }
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

        public void setOverlay(Overlay overlayWindow)
        {
            this.overlayWindow = overlayWindow;
        }

        public LeapListener()
        {

        }

        ~LeapListener()
        {

        }


    }
}*/
