using FirebirdSql.Data.FirebirdClient;

namespace KsaweryAPP;

public class FirebirdConnector
{
    private FbConnection _connection;

    public FirebirdConnector(string connectionString)
    {
        _connection = new FbConnection(connectionString);
    }
    /*SELECT GK.ID_GRUPAKONTRAH, GK.GRU_ID_GRUPAKONTRAH, GK.ID_RODZGRUPKTR, GK.KODGRUPY, GK.NAZWAGRUPY, GK.KODZLOZONY, GK.NAZWAZLOZONA, GK.ID_SLOWNIK, GK.AKTYWNY, SLOW.ID_SLOWNIK FROM GrupaKontrah GK INNER JOIN SLOWNIK SLOW ON (GK.ID_SLOWNIK = SLOW.ID_SLOWNIK)  WHERE (GK.ID_RodzGrupKtr=10002)*/
    public List<GrupaKontrah> GetGroupsKontrah()
    {
        List<GrupaKontrah> groups = new List<GrupaKontrah>();
        try
        {
            _connection.Open();
            using (var command = new FbCommand("SELECT GK.KODZLOZONY, GK.NAZWAZLOZONA FROM GrupaKontrah GK INNER JOIN SLOWNIK SLOW ON (GK.ID_SLOWNIK = SLOW.ID_SLOWNIK) WHERE (GK.ID_RodzGrupKtr=10002)", _connection))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        groups.Add(new GrupaKontrah{
                            KodZl = reader.GetInt32(0),
                            NazwaZl = reader.GetString(1)
                        });
                    }
                }
            }
            return groups;
        }
        catch (FbException ex)
        {
            Console.WriteLine(ex.Message);
            return groups;
        }
    }
}