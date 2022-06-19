using System.Collections.Generic;

namespace ClosedXML.Report.Tests.TestModels
{
    public class ParentsModel
    {
        public string Name { get; set; }
        public List<Parent> Parents { get; set; } = new();
    }

    public class Parent
    {
        public string Name { get; set; } = "Parent Name";
        public string ParentName => Name;

        public List<Child> Children { get; } = new()
        {
            new Child("Child 1"),
            new Child("Child 2"),
            new Child("Child 3"),
        };
    }

    public class Child
    {
        public string ChildName { get; }

        public Child(string childName)
        {
            ChildName = childName;
        }
    }

    public class Container
    {
        public List<Child> ItemsInContainer { get; } = new()
        {
            new Child("Item in container 1"),
            new Child("Item in container 2"),
            new Child("Item in container 3"),
        };
    }
}
