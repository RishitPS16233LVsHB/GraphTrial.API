namespace GraphTrial.API.Entities
{
    public class CreateTeam
    {
        public string TeamName { get; set; }
        public string TeamDescription { get; set; }
        public bool IsPrivate { get; set; }
        public string OwnerUserPrincipal { get; set; }
    }
}
