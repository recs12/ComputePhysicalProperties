open System
open SolidEdgeCommunity.Extensions
open System.Collections

[<STAThread>]
[<EntryPoint>]
let main argv =
    SolidEdgeCommunity.OleMessageFilter.Register()
    let application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true)
    let sheetMetalDocument = application.GetActiveDocument<SolidEdgePart.SheetMetalDocument>(false)

    let density: float = 1.0
    let accuracy = 0.05

    let density : double = 0.0
    let accuracy : double = 0.0

    //out
    let mutable volume  = 0.0
    let mutable area  = 0.0
    let mutable mass  = 0.0


    let mutable centerOfGravity = Array.CreateInstance(typeof<double>, 3)
    let mutable centerOfVolumne = Array.CreateInstance(typeof<double>, 3)
    let mutable globalMomentsOfInteria = Array.CreateInstance(typeof<double>, 6)     // Ixx, Iyy, Izz, Ixy, Ixz and Iyz
    let mutable principalMomentsOfInteria = Array.CreateInstance(typeof<double>, 3)  // Ixx, Iyy and Izz
    let mutable principalAxes = Array.CreateInstance(typeof<double>, 9)              // 3 axes x 3 coords
    let mutable radiiOfGyration = Array.CreateInstance(typeof<double>, 9)    // 3 axes x 3 coords
    let mutable relativeAccuracyAchieved = 0.0
    let mutable status = 0

    let models = sheetMetalDocument.Models
    let model = models.Item(1)

    // Compute the physical properties.
    model.ComputePhysicalProperties(
        Density = density,
        Accuracy = accuracy,
        Volume = &volume,
        Area = &area,
        Mass = &mass,
        CenterOfGravity= &centerOfGravity,
        CenterOfVolume= &centerOfVolumne,
        GlobalMomentsOfInteria= &globalMomentsOfInteria,
        PrincipalMomentsOfInteria= &principalMomentsOfInteria,
        PrincipalAxes= &principalAxes,
        RadiiOfGyration= &radiiOfGyration,
        RelativeAccuracyAchieved= &relativeAccuracyAchieved,
        Status= &status
        )

    printfn "ComputePhysicalProperties() results:"

    // Write results to screen.

    printfn "Density : %A" density
    printfn "Accuracy: %A" accuracy
    printfn "Volume  : %A" volume
    printfn "Area    : %A" area
    printfn "Mass    : %A" mass



    // Convert from System.Array to double[].  double[] is easier to work with.
    printfn "CenterOfGravity:"
    printfn "%A" centerOfGravity.Length

    let gravityCenter =
        centerOfGravity
        |> Seq.cast<double>
        |> Seq.toArray

    printfn "%A %A %A"  gravityCenter.[0] gravityCenter.[1]  gravityCenter.[2]


    SolidEdgeCommunity.OleMessageFilter.Unregister()
    Console.ReadKey() |> ignore
    0 // return an integer exit code

