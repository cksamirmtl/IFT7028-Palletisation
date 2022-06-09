using System;
using System.Activities.Expressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*Add from NuGet:
LinkToExcel v1.11.0
SimSharp v3.4.1*/

namespace IFT7028_Palletisation
{
    internal class Program
    {
        //******Paramètres******//
        private static string xlsxInputFilePathPanneaux = @"C:\Local\E2022\Conception et simulation des systèmes intelligents pour l'industrie 4.0 - (IFT-7028)\Projet\20220609\C-3312_DataSet.xlsm";
        private static string xlsxInputFilePathPalettes = @"C:\Local\E2022\Conception et simulation des systèmes intelligents pour l'industrie 4.0 - (IFT-7028)\Projet\20220609\Palette.xlsx";
        //private static string panneauxWorksheet = @"C-3312";
        //private static string palettesWorksheet = @"C-3312_palette";
        //private static int nbPalettesSurPlancher = 5;
        //Inter arrivé des paneaux entre 3 et 5 minutes
        private static readonly SimSharp.UniformTime PanneauxInterArrivalTime = SimSharp.Distributions.UNIF(new TimeSpan(0, 0, 0), new TimeSpan(0, 0, 1));
        //Temps traitement panneau 2,3,4 minutes triangulaire
        private static readonly SimSharp.TriangularTime PanneauxTraitementTime = SimSharp.Distributions.TRI(new TimeSpan(0, 2, 0), new TimeSpan(0, 4, 0), new TimeSpan(0, 3, 0));
        //Temps manutention lorsqu'une palette est terminée (entre 5 et 10 minutes)
        private static readonly SimSharp.UniformTime PaletteHandlingTime = SimSharp.Distributions.UNIF(new TimeSpan(0, 0, 0), new TimeSpan(0, 0, 1));
        //Simulation time in hours
        private static readonly TimeSpan SimTime = TimeSpan.FromHours(10000);
        //Initialiser avec des zéros l'ensemble de palettes simultannées
        //static List<PaletteStruct> palettes_encours1 = new List<PaletteStruct>()
        static List<object> palettes_encours1 = new List<object>()
        {
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
        };
        //******Paramètres******//


        IEnumerable<object> palettes_encours2 = (IEnumerable<object>)palettes_encours1;

        public struct PanneauStruct
        {
            public string No;
            public string Id;
            //public string Type;
            //public int Secteur;
            public int Num_palette;
        }

        public struct PaletteStruct
        {
            public int num_palette;
            public int quantity;
        }

        //PaletteStruct[] palettes_encours = new PaletteStruct[nbPalettesSurPlancher];

        DateTime simulationStartTime;

        delegate Func<object, bool> SelectorDelegate(object palletisationService);
        Func<object, bool> SelectorDelegateImp = delegate (object palletisationService)
        {
            if (palettes_encours1.Exists(x => ((PaletteStruct)x).num_palette == ((MyObject)palletisationService).panneau.Num_palette))
                return true;
            //if (palettes_encours1.Count < nbPalettesSurPlancher)
            //    return true;
            foreach (PaletteStruct palette_encours in palettes_encours1)
            {
                //si une palette est pleine, ou une palette fictive: on la remplace par la pelette qu'on a besoin pour notre panneau
                if ((palette_encours.num_palette == 0) || (all_palettes.First(palette => palette.num_palette == palette_encours.num_palette).quantity == palette_encours.quantity))
                {
                    return true;
                }
            }
            return false;
        };

        public struct MyObject
        {
            public PanneauStruct panneau;
            public SimSharp.ResourcePool palletisationService;
        }

        static IQueryable<PaletteStruct> all_palettes;
        static IQueryable<PanneauStruct> all_panneau;
        static List<PanneauStruct> tampon = new List<PanneauStruct>();
        //static IQueryable<PanneauStruct> tampon;

        private static object myLock = new object();
        static void Main(string[] args)
        {
            Console.WriteLine("Enter to start..");
            Console.ReadLine();
            new Program().SimulateSequentiel();
            //new Program().Simulate();
            Console.WriteLine("Done.");
            Console.ReadLine();
        }

        
        public void ReadData()
        {
            var excelFile = new LinqToExcel.ExcelQueryFactory(xlsxInputFilePathPanneaux);
            all_panneau =
                from row in excelFile.Worksheet(0) // (panneauxWorksheet)
                let item = new PanneauStruct()
                {
                    No = row["No"].Cast<string>(),
                    Id = row["Id"].Cast<string>(),
                    //Type = row["Type"].Cast<string>(),
                    //Secteur = row["Secteur"].Cast<int>(),
                    Num_palette = row["No_Palette"].Cast<int>(),
                }
                where item.Num_palette > 0
                select item;


            excelFile = new LinqToExcel.ExcelQueryFactory(xlsxInputFilePathPalettes);
            all_palettes =
                from row in excelFile.Worksheet(0) //(palettesWorksheet)
                let item = new PaletteStruct()
                {
                    num_palette = row["No"].Cast<int>(),
                    quantity = row["Qty"].Cast<int>(),
                }
                where item.num_palette > 0
                select item;

        }
        
        public void SimulateSequentiel()
        {
            ReadData();
            foreach (PanneauStruct panneau in all_panneau)
            {
                Console.WriteLine("Processing Panneau: " + panneau.No.ToString());
                //Chercehr si notre panneau a sa place
                //Si non traité, on insère dans le Tampon
                if (!traiterPanneau(panneau))
                {
                    tampon.Add(panneau);
                    //ImpressionTailleTampon();
                    //Console.Write("a");
                    ImpressionEtatPalettes();
                }

            }
        }

        //TODO: Optimisation ajouter pour palettes_encours1 un field qty max pour ne pas rechercher cette info à chaque fois dans all_palettes
        public bool traiterPanneau(PanneauStruct panneau, bool fromTampon = false)
        {

            bool traite = false;
            foreach (PaletteStruct palette_encours in palettes_encours1)
            {
                //On recherche une palette avec le meme numéro
                int indexPalettes_encours = palettes_encours1.FindIndex(x => ((PaletteStruct)x).num_palette == panneau.Num_palette);
                if (indexPalettes_encours != -1)
                {
                    //Trouvée. Est ce que pas pleine:
                    if (((PaletteStruct)palettes_encours1[indexPalettes_encours]).quantity < all_palettes.First(x => x.num_palette == ((PaletteStruct)palettes_encours1[indexPalettes_encours]).num_palette).quantity)
                    {
                        insertPanneauInPalette(panneau, indexPalettes_encours, fromTampon);

                        traite = true;

                        break;
                    }
                    //PaletteStruct pl = ((PaletteStruct)palettes_encours1[i]);
                }

                //Si pas sorti avec le break, on recherche une palette libre (0,0)
                indexPalettes_encours = palettes_encours1.FindIndex(x => ((PaletteStruct)x).num_palette == 0);
                if (indexPalettes_encours != -1)
                {
                    insertPanneauInPalette(panneau, indexPalettes_encours, fromTampon);
                    traite = true;
                    break;
                }


            }
            return traite;
        }

        public void ImpressionTailleTampon()
        {
            Console.WriteLine("                                                                                              Taille Tampon: " + tampon.Count().ToString());
        }
        public void ImpressionEtatPalettes()
        {
            //Impression
            Console.Write("                                  Etat palettes: ");
            foreach (PaletteStruct palette_encours in palettes_encours1)
            {
                Console.Write(palette_encours.num_palette.ToString() + ": " + palette_encours.quantity.ToString() + "/");
                int qtyTotal = 0;
                try
                {
                    qtyTotal = all_palettes.First(x => x.num_palette == palette_encours.num_palette).quantity;
                }
                catch (Exception ex) { }
                Console.Write(qtyTotal.ToString() + " - ");
            }
            Console.Write(" Taille Tampon: " + tampon.Count().ToString());
            Console.WriteLine();
        }

        public void insertPanneauInPalette(PanneauStruct panneau, int indexPalettes_encours, bool fromTampon = false)
        {

            //TODO: ajouter "Provenance" dans la fonction d'insertion dans la palette afin de mettre à jour le Tampon instantannément


            PaletteStruct pl = ((PaletteStruct)palettes_encours1[indexPalettes_encours]);
            pl.quantity++;
            pl.num_palette = panneau.Num_palette;
            palettes_encours1[indexPalettes_encours] = pl;
            if (fromTampon)
            {
                tampon.Remove(panneau);
            }
            //Console.Write("b");

            ImpressionEtatPalettes();

            //Si palette full, on la sort
            if (((PaletteStruct)palettes_encours1[indexPalettes_encours]).quantity == all_palettes.First(x => x.num_palette == ((PaletteStruct)palettes_encours1[indexPalettes_encours]).num_palette).quantity)
            {
                //TODO: Ajouter Temps de latence
                //La palette qui remplace sera la première dans le Tampon et on fait le tour du Tampon. Si vide on ajoute une palette libre (0,0)
                if (tampon.Count > 0)
                {
                    PanneauStruct pnFromTampon = tampon[0];
                    //tampon.Remove(pnFromTampon);//deplacé dans la méthode
                    //Console.Write("1");
                    //ImpressionTailleTampon();
                    //Console.Write("c");
                    //ImpressionEtatPalettes();
                    //on met qté à 0 avant d'insérer car on commence la palette
                    PaletteStruct plInit = ((PaletteStruct)palettes_encours1[indexPalettes_encours]);
                    pl.quantity = 0;
                    pl.num_palette = pnFromTampon.Num_palette;//Pas nécessaire car faite à l'appel de insertPanneauInPalette
                    palettes_encours1[indexPalettes_encours] = pl;
                    
                    insertPanneauInPalette(pnFromTampon, indexPalettes_encours, true    );
                    //faire le tour du Tampon (sans oublier que si ça s'insert pas ne pas mettre dans Tampon mais plutot supprimer du Tampon
                    

                    var tamponCopy = new List<PanneauStruct>(tampon);
                    foreach (PanneauStruct panneauDansTampon in tamponCopy)
                    {

                        //tkhaf tu l'insere deux fois, car dès qu'une palete est pleine tu prend celui de la position 0
                        //donc avant de traiter vérifier qu'il est là
                        if ((tampon.IndexOf(panneauDansTampon) != -1) && traiterPanneau(panneauDansTampon, true))
                        {
                            //TODO: ne pas remover par index mais par ID du panneau
                            //if (tampon.Remove(panneauDansTampon))//déplacé dans méthode
                            //{
                            //    //hadi sa3at ydirha retard
                            //    //Console.Write("2");
                            //    //ImpressionTailleTampon();
                            //    Console.Write("d");
                            //    ImpressionEtatPalettes();
                            //}

                        }
                    }
                }
                else //sinon ça sera une palette libre (0,0)
                {
                    PaletteStruct paletteLibre = new PaletteStruct();
                    paletteLibre.num_palette = 0;
                    paletteLibre.quantity = 0;
                    palettes_encours1[indexPalettes_encours] = paletteLibre;
                }
            }
        }

        //Nb dans Tampon
        //Temps mloyen pour remplir Palette
        //Temps total de la paletisation

        public void Simulate()
        {
            int RandomSeed = 77;
            var env = new SimSharp.Simulation(RandomSeed);
            //palettes_encours1.Where(p => ((PaletteStruct)p).num_palette == 2);
            //var palletisationService = new SimSharp.Resource(env, nbPalettesSurPlancher)
            //var palletisationService = new SimSharp.ResourcePool(env, (IEnumerable<object>)(System.Collections.IEnumerable)palettes_encours)
            var palletisationService = new SimSharp.ResourcePool(env, (IEnumerable<object>)palettes_encours2)
            {
                //des metriques
                QueueLength = new SimSharp.TimeSeriesMonitor(env, name: "Queue", collect: true),
                //
            };

            ReadData();


            env.Process(PanneauxGenerator(env, palletisationService));
            /***RUN***/
            simulationStartTime = env.Now;
            env.Run(SimTime);
            /*********/
        }

        private IEnumerable<SimSharp.Event> PanneauxGenerator(SimSharp.Simulation env, SimSharp.ResourcePool palletisationService)
        {
            foreach (PanneauStruct panneau in all_panneau)
            {
                //Console.WriteLine("PanneauNo {0}:", panneau.No);
                //yield return env.Timeout(PanneauxInterArrivalTime);
                yield return env.Timeout(new TimeSpan(0));
                env.Process(PanneauProcess(panneau, env, palletisationService));
            }
        }

        //Func<object, bool> selector = palettes_encours => ((IEnumerable<object>)(System.Collections.IEnumerable)palettes_encours).;
        //Func<object, bool> selector = x => ((IEnumerable<object>)(System.Collections.IEnumerable)x).;
        //Func<object, bool> selector = x => ((PaletteStruct)x).num_palette == 2;

        private IEnumerable<SimSharp.Event> PanneauProcess(PanneauStruct panneau, SimSharp.Simulation env, SimSharp.ResourcePool palletisationService)
        {
            DateTime arrivingTime = env.Now;
            env.Log("Processing Panneau: " + panneau.No.ToString());

            //Request access to palletisation service
            MyObject myObject = new MyObject();
            myObject.panneau = panneau;
            myObject.palletisationService = palletisationService;
            //SelectorDelegate selector = SelectorDelegateImp;
            using (var req = palletisationService.Request(new Func<object, bool>((o) => SelectorDelegateImp(myObject))))
            {
                //Request access to palletisation service
                Again:
                yield return req;

                //Checker si la palette pour notre panneau est disponible, sinon mettre le panneau dans le Tampon
                //Ce check doit être fait après l'autorisation d'accès au service, car la liste des palettes en cours peut changer

                //mettre à jour la liste palettes_encours
                //lock (myLock)
                //{
                    if (palettes_encours1.Exists(x => ((PaletteStruct)x).num_palette == panneau.Num_palette))
                    {
                        int i = palettes_encours1.FindIndex(x => ((PaletteStruct)x).num_palette == panneau.Num_palette);
                        PaletteStruct pl = ((PaletteStruct)palettes_encours1[i]);
                        pl.quantity++;
                        if (pl.quantity <= all_palettes.First(x => x.num_palette == pl.num_palette).quantity)
                            palettes_encours1[i] = pl;
                        else
                            goto Again;
                    }
                    else
                    {
                        foreach (PaletteStruct palette_encours in palettes_encours1.ToList())
                        {
                            //si une palette est pleine, ou une palette fictive: on la remplace par la pelette qu'on a besoin pour notre panneau
                            if ((palette_encours.num_palette == 0) || (all_palettes.First(palette => palette.num_palette == palette_encours.num_palette).quantity == palette_encours.quantity))
                            {
                                if (!(palette_encours.num_palette == 0)) 
                                {
                                    //yield return env.Timeout(PaletteHandlingTime);
                                }
                                //else break;
                                PaletteStruct pl = all_palettes.First(palette => palette.num_palette == panneau.Num_palette);
                                pl.quantity = 1;
                                int i = palettes_encours1.FindIndex(x => ((PaletteStruct)x).num_palette == palette_encours.num_palette);
                                if (i != -1)
                                    palettes_encours1[i] = pl;
                                else 
                                    goto Again;
                                break;
                            }
                        }
                    }
                //}
                Console.Write("Etat palettes: ");
                foreach (PaletteStruct palette_encours in palettes_encours1)
                {
                    Console.Write(palette_encours.num_palette.ToString() + ": " + palette_encours.quantity.ToString() + "/");
                    int qtyTotal = 0;
                    try
                    {
                        qtyTotal = all_palettes.First(x => x.num_palette == palette_encours.num_palette).quantity;
                    }
                    catch (Exception ex) { }
                    Console.Write(qtyTotal.ToString() + " - ");
                }
                Console.WriteLine();
                // Traitement panneau
                //yield return env.Timeout(PanneauxTraitementTime);

            }
        }
    }
}
