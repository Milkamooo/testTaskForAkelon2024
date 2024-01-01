using System;
namespace task4.Models
{
	public class Order : Base
	{
		public int Number { get; set; }
		public int Amount { get; set; }
		public DateTime DateOfCreate { get; set; }
		public int ProductId { get; set; }
		public int ClientId { get; set; }
	}
}
