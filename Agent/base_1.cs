
namespace Agent
{
    class datebase
    {
        public string connection = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\Diplom\proga\Agent\Agent\Agent.mdf;Integrated Security=True;Connect Timeout=30";
        //public string connection = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Agent.mdf;Integrated Security=True;Connect Timeout=30";

        //public string Organization = "SELECT idorganization, name as Название, unp as УНП, address as Адрес, faks as Факс, email as Почта FROM Organization ";

        //public string Orderr = "SELECT Orderr.Idorder, Orderr.idclient, (Client.f+' '+Client.i+' '+Client.o) as Клиент,Orderr.idseller,  (Seller.f + ' ' + Seller.i + ' ' + Seller.o) as Сотрудник, date as Дата, vidoplat as [Вид оплаты],  status as Статус, Sum(CASE WHEN(Orderr.Idorder= Sostav.idorder) THEN(count* cost) else null end) as Стоимость FROM Sostav,Orderr,Client,Seller,Clothes where Orderr.idclient=Client.Idclienta and Orderr.idseller= Seller.Idseller and Clothes.Idclothes=Sostav.idclothes Group by Orderr.Idorder, Orderr.idclient, (Client.f+' '+Client.i+' '+Client.o) ,Orderr.idseller, (Seller.f+' '+Seller.i+' '+Seller.o) ,date , vidoplat , status";
        //public string Clothes = "SELECT Idclothes, name as 'Наименование',pol as Пол,size as Размер, art as Артикул, ost as Количество, cost as [Цена за шт.] FROM Clothes ";

        //public string Client = "SELECT idclienta, f as Фамилия, i as Имя, o as Отчество, address as Адрес, phone as Телефон, email as Почта FROM Client ";

        //public string Seller = "SELECT Idseller, f as Фамилия, i as Имя, o as Отчество, idorg, name as Организация FROM Seller inner join Organization on Seller.idorg=Organization.Idorganization ";

    }
}
