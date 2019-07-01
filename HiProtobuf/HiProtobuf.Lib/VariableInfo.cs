namespace HiProtobuf.Lib
{
    public class VariableInfo
    {
        public string Type { get; private set; }
        public string Name { get; private set; }

        public VariableInfo(string type, string name)
        {
            this.Type = type;
            this.Name = name;
        }
    }
}
