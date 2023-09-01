namespace Svero.HelloExcel
{
    public class JournalTransaction
    {
        public long Id { get; set; }

        public DateTime PlanDate { get; set; }

        public DateTime? ActualDate { get; set; }

        public string? Description { get; set; }

        public decimal Value { get; set; }

        public override string ToString()
        {
            return $"JournalTransaction(Id: {Id}, PlanDate: {PlanDate.ToShortDateString()}, " + 
                $"ActualDate: {ActualDate?.ToShortDateString()}, Description: {Description}, Value: {Value})";
        }
    }
}