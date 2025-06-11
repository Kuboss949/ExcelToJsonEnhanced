namespace KsaweryAPP;

public class GrupaKontrah
{
    public int KodZl { get; set; }
    public string NazwaZl { get; set; }
    public override string ToString()
    {
        return KodZl.ToString() + " " + NazwaZl.ToString();
    }
}