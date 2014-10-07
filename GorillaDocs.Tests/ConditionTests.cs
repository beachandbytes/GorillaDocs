using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace GorillaDocs.Tests
{
    [TestFixture]
    public class ConditionTests
    {
        [Test]
        public void Equals()
        {
            var result = new OptionalCondition("Grade = 5", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void NotEquals()
        {
            var result = new OptionalCondition("Grade != 4", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void NotEquals_with_different_operator()
        {
            var result = new OptionalCondition("Grade <> 6", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void LessThan()
        {
            var result = new OptionalCondition("Grade < 7", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void LessThanOrEqual()
        {
            var result = new OptionalCondition("Grade <= 5", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void GreaterThan()
        {
            var result = new OptionalCondition("Grade > 4", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void GreaterThanOrEqual()
        {
            var result = new OptionalCondition("Grade >= 5", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }
        [Test]
        public void Or()
        {
            var result = new OptionalCondition("Grade = 5 or Grade = 6", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }

        [Test]
        public void Or1()
        {
            var result = new OptionalCondition("Grade = 5 OR Grade = 6 || Grade = 7", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }

        [Test]
        public void Or_when_false()
        {
            var result = new OptionalCondition("Grade = 7 OR Grade = 8 || Grade = 9", GetGradeData()).Evaluate();
            Assert.That(result == false);
        }

        [Test]
        public void Null_Manager()
        {
            var result = new OptionalCondition("Manager != Test", GetGradeData()).Evaluate();
            Assert.That(result == true);
        }

        static XDocument GetGradeData() { return XDocument.Parse("<root><Grade>5</Grade><Manager/></root>"); }
    }
}
