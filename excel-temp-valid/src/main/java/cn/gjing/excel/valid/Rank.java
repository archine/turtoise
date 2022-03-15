package cn.gjing.excel.valid;

/**
 * Error box grade
 *
 * @author Gjing
 **/
public enum Rank {
    /**
     * Tip rank
     */
    WARNING(1), STOP(0), INFO(2);
    private final int rank;

    Rank(int type) {
        this.rank = type;
    }

    public int getRank() {
        return rank;
    }
}
