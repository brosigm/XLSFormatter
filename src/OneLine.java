public class OneLine {
    public String nev;
    public double x, y;
    public String value;

    public OneLine(){
        nev = "";
        x = 0;
        y = 0;
        value = "";
    }

    public double getDistanceFrom(OneLine other){
        return Math.sqrt(Math.pow(x - other.x, 2) + Math.pow(y - other.y, 2));
    }
}
