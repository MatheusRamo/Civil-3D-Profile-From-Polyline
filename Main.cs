using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System.Linq;


namespace CreateProfileFromPolyline
{
    public class Main 
    {

        [CommandMethod("CREATE_PROFILE_FROM_POLYLINE")]
        public void CreateProfileFromPolyline()
        {
            DocumentCollection docCol = Application.DocumentManager;
            Database database = docCol.MdiActiveDocument.Database;
            Editor editor = docCol.MdiActiveDocument.Editor;
            CivilDocument doc = CivilApplication.ActiveDocument;


            using (Transaction tx = database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Select Polyline
                    PromptEntityOptions promptPolyline = new PromptEntityOptions("\nSelect a Polyline: ");
                    promptPolyline.SetRejectMessage("\n Polyline not Select");
                    promptPolyline.AddAllowedClass(typeof(Polyline), true);
                    PromptEntityResult entityPolyline = editor.GetEntity(promptPolyline);
                    if (entityPolyline.Status != PromptStatus.OK) return;


                    // Select Profile View 
                    PromptEntityOptions promptProfileView = new PromptEntityOptions("\nSelect a profile view: ");
                    promptProfileView.SetRejectMessage("\nProfileView not select");
                    promptProfileView.AddAllowedClass(typeof(ProfileView), true);
                    PromptEntityResult entityProfileView = editor.GetEntity(promptProfileView);
                    if (entityProfileView.Status != PromptStatus.OK) return;



                    ProfileView profileView = tx.GetObject(entityProfileView.ObjectId, OpenMode.ForWrite) as ProfileView;
                    double x = 0.0;
                    double y = 0.0;
                    if (profileView.ElevationRangeMode == ElevationRangeType.Automatic)
                    {
                        profileView.ElevationRangeMode = ElevationRangeType.UserSpecified;
                        profileView.FindXYAtStationAndElevation(profileView.StationStart, profileView.ElevationMin, ref x, ref y);
                    }
                    else
                        profileView.FindXYAtStationAndElevation(profileView.StationStart, profileView.ElevationMin, ref x, ref y);

                    ProfileViewStyle profileViewStyle = tx.GetObject(profileView.StyleId, OpenMode.ForRead) as ProfileViewStyle;

                    ObjectId layerId = (tx.GetObject(profileView.AlignmentId, OpenMode.ForRead) as Alignment).LayerId;

                    ObjectId profileStyleId = doc.Styles.ProfileStyles.FirstOrDefault();

                    ObjectId profileLabelSetStylesId = doc.Styles.LabelSetStyles.ProfileLabelSetStyles.FirstOrDefault();

                    ObjectId profByLayout = Profile.CreateByLayout("New Profile", profileView.AlignmentId, layerId, profileStyleId, profileLabelSetStylesId);

                    Profile profile = tx.GetObject(profByLayout, OpenMode.ForWrite) as Profile;

                    BlockTableRecord blockTableRecord = tx.GetObject(database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    ObjectId polylineObjId = entityPolyline.ObjectId;

                    Polyline polyline = tx.GetObject(polylineObjId, OpenMode.ForWrite, false) as Polyline;


                    double intervalMajorTick = profileViewStyle.BottomAxis.MajorTickStyle.Interval;
                    double gridPadding = profileViewStyle.GridStyle.GridPaddingLeft * intervalMajorTick;

                    //Invert the Polyline if the start coordinate is greater than the end coordinate
                    if (polyline != null && (polyline.StartPoint.X > polyline.EndPoint.X))
                    {
                        polyline.ReverseCurve();

                    }

                    if (polyline != null)
                    {
                        int numOfVert = polyline.NumberOfVertices - 1;
                        Point2d startSegment;
                        Point2d endSegment;
                        Point2d sampleSegment;
                        Point2d coodStartSeg;
                        Point2d coodEndSeg;
                        Point2d coodSampleSegment;

                        for (int i = 0; i < numOfVert; i++)
                        {
                            switch (polyline.GetSegmentType(i))
                            {
                                case SegmentType.Line:
                                    LineSegment2d lineSegment2dAt = polyline.GetLineSegment2dAt(i);

                                    startSegment = lineSegment2dAt.StartPoint;
                                    double difStartSegment_X = (startSegment.X - gridPadding) - x;
                                    double difStartSegment_Y = (startSegment.Y - y) / profileViewStyle.GraphStyle.VerticalExaggeration + profileView.ElevationMin;
                                    coodStartSeg = new Point2d(difStartSegment_X, difStartSegment_Y);
                             

                                    endSegment = lineSegment2dAt.EndPoint;
                                    double difEndSegment_X = (endSegment.X - gridPadding) - x;
                                    double difEndSegment_Y = (endSegment.Y - y) / profileViewStyle.GraphStyle.VerticalExaggeration + profileView.ElevationMin;
                                    coodEndSeg = new Point2d(difEndSegment_X, difEndSegment_Y);

                                
                                    profile.Entities.AddFixedTangent(coodStartSeg, coodEndSeg);
                                    break;
                                case SegmentType.Arc:
                                    CircularArc2d arcSegment2dAt = polyline.GetArcSegment2dAt(i);

                                    startSegment = arcSegment2dAt.StartPoint;
                                    double difStartSegmentArc_x = (startSegment.X - gridPadding) - x;
                                    double difStartSegmentArc_y = (startSegment.Y - y) / profileViewStyle.GraphStyle.VerticalExaggeration + profileView.ElevationMin;
                                    coodStartSeg = new Point2d(difStartSegmentArc_x, difStartSegmentArc_y);


                                    endSegment = arcSegment2dAt.EndPoint;
                                    double difEndSegmentArc_x = endSegment.X - gridPadding - x;
                                    double difEndSegmentArc_y = (endSegment.Y - y) / profileViewStyle.GraphStyle.VerticalExaggeration + profileView.ElevationMin;
                                    coodEndSeg = new Point2d(difEndSegmentArc_x, difEndSegmentArc_x);

                                  
                                    sampleSegment = arcSegment2dAt.GetSamplePoints(11)[5];
                                    double difMidSegmentArc_x = (sampleSegment.X - gridPadding) - x;
                                    double difMidSegmentArc_y = (sampleSegment.Y - y) / profileViewStyle.GraphStyle.VerticalExaggeration + profileView.ElevationMin;
                                    coodSampleSegment = new Point2d(difMidSegmentArc_x, difMidSegmentArc_y);

                                    profile.Entities.AddFixedSymmetricParabolaByThreePoints(coodStartSeg, coodSampleSegment, coodEndSeg);
                                    break;
                                case SegmentType.Coincident:
                                    break;
                                case SegmentType.Point:
                                    break;
                                case SegmentType.Empty:
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    editor.WriteMessage("\n" + ex.Message);
                }
                tx.Commit();
            }
        }
    }
}