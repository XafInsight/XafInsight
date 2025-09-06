namespace xafplugin.Modules
{
    public enum EJoinType
    {
        Inner,         // INNER JOIN: matching records in both tables
        LeftOuter,     // LEFT (OUTER) JOIN: all from left, matched from right
        RightOuter,    // RIGHT (OUTER) JOIN: all from right, matched from left
        FullOuter      // FULL (OUTER) JOIN: all records from both sides
    }
}
