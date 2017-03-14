/*
    Permission is granted to copy, distribute and/or modify this document
    under the terms of the GNU Free Documentation License, Version 1.3
    or any later version published by the Free Software Foundation;
    with no Invariant Sections, no Front-Cover Texts, and no Back-Cover Texts.
    A copy of the license is included in the section entitled "GNU
    Free Documentation License".
*/

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Collections.Generic;
using Microsoft.Kinect;
using System.Linq;
using MIDIWrapper;


namespace MyKinect
{
    public partial class MainWindow : Window
    {
        //Instantiate the Kinect runtime. Required to initialize the device.
        //IMPORTANT NOTE: You can pass the device ID here, in case more than one Kinect device is connected.

 


        KinectSensor sensor = KinectSensor.KinectSensors[0];
        byte[] pixelData;
        Body[] skeletons;
        int[] zaehler;
        int[] SumNoteLeft;
        int[] SumNoteRight;
        byte[] LastNoteLeft;
        byte[] LastNoteRight;

        Instrument MyInstrument;

        public MainWindow()
        {
            InitializeComponent();

            //Runtime initialization is handled when the window is opened. When the window
            //is closed, the runtime MUST be unitialized.
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
            this.Unloaded += new RoutedEventHandler(MainWindow_Unloaded);

            sensor.ColorStream.Enable();
            sensor.SkeletonStream.Enable();
            MyInstrument = new Instrument();
            string[] OutMidiNames = Instrument.OutDeviceNames();
            zaehler = new int[4];
            SumNoteLeft = new int[4];
            SumNoteRight = new int[4];
            LastNoteLeft = new byte[4];
            LastNoteRight = new byte[4];

            MyInstrument.OutputDeviceName = OutMidiNames[1];
            MyInstrument.Open();
        }

        void runtime_SkeletonFrameReady(object sender, BodyFrameArrivedEventArgs e)
        {
            bool receivedData = false;

            using (BodyFrame SFrame = e.FrameReference.AcquireFrame())
            {
                if (SFrame == null)
                {
                    // The image processing took too long. More than 2 frames behind.
                }
                else
                {
                    skeletons = new Body[SFrame.SkeletonArrayLength];
                    SFrame.CopySkeletonDataTo(skeletons);
                    receivedData = true;
                }
            }

            if (receivedData)
            {

                IEnumerable<Body> sel = (from s in skeletons
                                             where s.TrackingState == SkeletonTrackingState.Tracked
                                             select s);
                int num = 0;
                foreach (Body currentSkeleton in sel)
                {
                    if (currentSkeleton != null)
                    {
                        processSkeleton(num, currentSkeleton);
                        Console.WriteLine(num);
                    }
                    else
                    {
                        silenceSkeleton(num);
                    }
                    num++;
                }
                if (num == 0)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        silenceSkeleton(i);
                    }
                }
            }
            else
            {
                for (int i = 0; i < 4; i++)
                {
                    silenceSkeleton(i);
                }
            }
        }


        void processSkeleton(int num, Body skel)
        {
            SetEllipsePosition(head, skel.Joints[JointType.Head]);
            SetEllipsePosition(leftHand, skel.Joints[JointType.HandLeft]);
            SetEllipsePosition(rightHand, skel.Joints[JointType.HandRight]);


            //   DrawTrackedSkeletonJoints(currentSkeleton);
            byte NewNoteLeft = (byte)(skel.Joints[JointType.HandLeft].Position.Y * 100 + 75);
            byte NewNoteRight = (byte)(skel.Joints[JointType.HandRight].Position.Y * 100 + 75);
            if ((NewNoteLeft < 40) || (NewNoteLeft > 200))
                NewNoteLeft = 0;
            if ((NewNoteRight < 40) || (NewNoteRight > 200))   // Werte zwischen 40 und 200
                NewNoteRight = 0;
            byte NewNoteDepth = (byte)((byte)(skel.Joints[JointType.HandLeft].Position.Z)*7);
            if (NewNoteDepth > 28)
                NewNoteDepth = 28;   // 4 quintas

            Console.WriteLine("Depth : {0}", NewNoteDepth);
            /*    
                Console.WriteLine("Left :   {0:N}", NewNoteLeft);
                  Console.WriteLine("Right :  {0:N}", NewNoteRight);
                
             */
            zaehler[num]++;
          
            SumNoteLeft[num] += NewNoteLeft;
            SumNoteRight[num] += NewNoteRight;
            if (zaehler[num] > 2)
            {

                byte noteleft = (byte)(SumNoteLeft[num] / 2);
                byte noteright = (byte)(SumNoteRight[num] / 2);
                byte chanNr = (byte)(num * 2);   // you need 2 midi channels for 1 skeleton
                noteleft = MyInstrument.GetMidiNote(noteleft);
                noteright = MyInstrument.GetMidiNote(noteright);
                if (noteleft > 0x00)
                    noteleft = (byte)(noteleft - NewNoteDepth);
                if (noteright > 0x00)
                    noteright = (byte)(noteright - NewNoteDepth);
  
                MyInstrument.TranslateToNote(chanNr, noteleft, LastNoteLeft[num], noteright, LastNoteRight[num]);
                if (noteleft > 0)
                    LastNoteLeft[num] = noteleft;
                if (noteright > 0)
                    LastNoteRight[num] = noteright;
                //             MyInstrument.TranslateToNote(NewNoteLeft, NewNoteRight);
                zaehler[num] = 0;
                SumNoteLeft[num] = 0;
                SumNoteRight[num] = 0;
            }
        }

        void silenceSkeleton(int num)
        {
            SetEllipsePosition(head, new Joint());
            SetEllipsePosition(leftHand,new Joint());
            SetEllipsePosition(rightHand, new Joint());

            byte noteleft = 0;
            byte noteright = 0;
            byte chanNr = (byte)(num * 2);   // you need 2 midi channels for 1 skeleton
           
            MyInstrument.TranslateToNote(chanNr, noteleft, LastNoteLeft[num], noteright, LastNoteRight[num]);
        }



        //This method is used to position the ellipses on the canvas
        //according to correct movements of the tracked joints.

        //IMPORTANT NOTE: Code for vector scaling was imported from the Coding4Fun Kinect Toolkit
        //available here: http://c4fkinect.codeplex.com/
        //I only used this part to avoid adding an extra reference.
        private void SetEllipsePosition(Ellipse ellipse, Joint joint)
        {
            /*
            Microsoft.Kinect.SkeletonPoint vector = new Microsoft.Kinect.SkeletonPoint();
            vector.X = ScaleVector(640, joint.Position.X);
            vector.Y = ScaleVector(480, -joint.Position.Y);
            vector.Z = joint.Position.Z;


            Joint updatedJoint = new Joint();
            updatedJoint = joint;
            updatedJoint.TrackingState = JointTrackingState.Tracked;
            updatedJoint.Position = vector;

            Canvas.SetLeft(ellipse, updatedJoint.Position.X);
            Canvas.SetTop(ellipse, updatedJoint.Position.Y);
  */
            ColorImagePoint point = sensor.MapSkeletonPointToColor(joint.Position, ColorImageFormat.RgbResolution640x480Fps30);
            double sizeX = canvas.ActualWidth;
            double sizeY = canvas.ActualHeight;
            point.X = (int)((point.X / 640.0) * sizeX);
            point.Y = (int)((point.Y / 480.0) * sizeY);

            Canvas.SetLeft(ellipse, point.X);
            Canvas.SetTop(ellipse, point.Y);
 
        }

/*
        private float ScaleVector(int length, float position)
        {
            float value = ((((float)(length/2)) * position) + (length / 2))*(1.15f);
            if (value > length)
            {
                return (float)length;
            }
            if (value < 0f)
            {
                return 0f;
            }
            return value;
        }
*/
        void MainWindow_Unloaded(object sender, RoutedEventArgs e)
        {
            sensor.Stop();
            MyInstrument.Close();
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            sensor.SkeletonFrameReady += runtime_SkeletonFrameReady;
            sensor.ColorFrameReady += runtime_VideoFrameReady;
            sensor.Start();
        }

        void runtime_VideoFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            bool receivedData = false;

            using (ColorImageFrame CFrame = e.OpenColorImageFrame())
            {
                if (CFrame == null)
                {
                    // The image processing took too long. More than 2 frames behind.
                }
                else
                {
                    pixelData = new byte[CFrame.PixelDataLength];
                    CFrame.CopyPixelDataTo(pixelData);
                    receivedData = true;
                }
            }

            if (receivedData)
            {
                BitmapSource source = BitmapSource.Create(640, 480, 96, 96,
                        PixelFormats.Bgr32, null, pixelData, 640 * 4);

                videoImage.Source = source;
            }
        }
    }
}
