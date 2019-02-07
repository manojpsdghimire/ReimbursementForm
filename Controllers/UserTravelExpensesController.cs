using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using ReimbursementVoucherForm.Models;

namespace ReimbursementVoucherForm.Controllers
{
    public class UserTravelExpensesController : Controller
    {
        private WCAPPSEntities db = new WCAPPSEntities();

        // GET: UserTravelExpenses
        public ActionResult Index()
        {
            var userTravelExpenses = db.TRF_UserTravelExpense.Include(u => u.TRF_File).Include(u => u.TRF_TravelAddress).Include(u => u.TRF_TravelExpense).Include(u => u.TRF_User);
            return View(userTravelExpenses.ToList());
        }

       
        // GET: UserTravelExpenses/Create
        public ActionResult Create()
        {
            ViewBag.FileID = new SelectList(db.TRF_File, "FileID", "FileName");
            ViewBag.TravelID = new SelectList(db.TRF_TravelAddress, "TravelID", "Address");
            ViewBag.ExpenseID = db.TRF_TravelExpense.ToList();
            ViewBag.UserID = new SelectList(db.TRF_User, "UserID", "LastName");
            return View();
        }

        // POST: UserTravelExpenses/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Amount,SubmittedDate,UserID,TravelID,ExpenseID,FileID")] List<TRF_UserTravelExpense> userTravelExpense, [Bind(Include = "UserID,LastName,FirstName,Email,UserType,Department,SupervisorName,AccountCode,TravelID,notescomments")] TRF_User user, [Bind(Include = "TravelID,Address,City,State,Country,ZipCode,DepartureDate,ArrivalDate")] TRF_TravelAddress travelAddress, List<HttpPostedFileBase> upload)
        {
            //working with files


            if (ModelState.IsValid)
            {
                /*************************************************/

                user.UserID = Guid.NewGuid();
                travelAddress.TravelID = Guid.NewGuid();
                DateTime Datesubmitted = System.DateTime.Now;
                Guid travelexpenseID = Guid.NewGuid();
                /*************************************************/
                //assigning values to the foreignkey
                travelAddress.TravelID = Guid.NewGuid();
                user.TravelID = travelAddress.TravelID;
                /*************************************************/
                //save to user and travel database 
                db.TRF_TravelAddress.Add(travelAddress);
                db.TRF_User.Add(user);
                //going through the travel list 
                int i = 0;          //variable for file index

                foreach (var item in userTravelExpense)
                {
                    item.ID = travelexpenseID;
                    item.TravelID = travelAddress.TravelID;
                    item.UserID = user.UserID;
                    item.SubmittedDate = Datesubmitted;

                    if (db.TRF_UserTravelExpense.Any(o => o.ExpenseID == item.ExpenseID && o.TravelID == item.TravelID && o.TravelID == item.TravelID && o.UserID == item.UserID && o.ID == item.ID))
                    {

                    }
                    else
                    {
                        if (item.Amount > 0.00)
                        {
                            if (upload[i] != null && upload[i].ContentLength > 0)
                            {
                                var avatar = new TRF_File
                                {
                                    FileName = System.IO.Path.GetFileName(upload[i].FileName),
                                    ContentType = upload[i].ContentType

                                };

                                using (var reader = new System.IO.BinaryReader(upload[i].InputStream))
                                {
                                    avatar.Content = reader.ReadBytes(upload[i].ContentLength);
                                    avatar.UserID = item.UserID;
                                    avatar.ExpenseID = item.ExpenseID;
                                    avatar.TravelID = travelAddress.TravelID;
                                    avatar.FileID = Guid.NewGuid();
                                    item.FileID = avatar.FileID;
                                    avatar.SubmittedDate = item.SubmittedDate;
                                }
                                db.TRF_File.Add(avatar);

                            }
                            db.TRF_UserTravelExpense.Add(item);
                        }

                        i++;
                        db.SaveChanges();
                    }

                }


                //sending email
                Sending_mail(user.UserID, travelAddress.TravelID, travelexpenseID);
                return RedirectToAction("Index");
            }

            ViewBag.FileID = new SelectList(db.TRF_File, "FileID", "FileName", userTravelExpense.Select(u => u.FileID));
            ViewBag.TravelID = new SelectList(db.TRF_TravelAddress, "TravelID", "Address", userTravelExpense.Select(u => u.TravelID));
            ViewBag.ExpenseID = db.TRF_TravelExpense.ToList();
            ViewBag.UserID = new SelectList(db.TRF_User, "UserID", "LastName", userTravelExpense.Select(u => u.UserID));
            return View();
        }

        [HttpPost]
        public void Sending_mail(Guid UID, Guid TID, Guid UTID)
        {
            if (db.TRF_UserTravelExpense.Any(o => o.UserID == UID && o.TravelID == TID && o.ID == UTID))
            {
                TRF_User users = db.TRF_User.Find(UID);

                if (users.TravelID == TID)
                {
                    //files in the server
                    var dir = Server.MapPath("~/ReimbursementPDF");
                    if (!System.IO.Directory.Exists(dir))
                        System.IO.Directory.CreateDirectory(dir);
                    var OutputPath = dir + "\\ReimbursementVoucher-" + users.FirstName + users.LastName + "-" + DateTime.Now.ToString("yyyyMMdd") + ".pdf";

                    //sending email
                    MailAddress addressFrom = new MailAddress(users.Email);
                    MailAddress addressTo = new MailAddress("cinnamon.brown@westminster-mo.edu");

                    //using (var client = new SmtpClient())
                    using (var mm = new MailMessage(addressFrom, addressTo))
                    {
                        SmtpClient smtp = new SmtpClient();
                        smtp.Host = System.Configuration.ConfigurationManager.AppSettings["Host"];
                        mm.Subject = "Reimbursement Voucher for " + users.FirstName + " " + users.LastName;
                        mm.Body = CreateBody(UID, TID, UTID);
                        mm.IsBodyHtml = true;
                        // mm.To.Add("matthew.vore@westminster-mo.edu");
                        // mm.To.Add("cinnamon.brown@westminster-mo.edu");
                        mm.To.Add(users.Email);

                        //sending files as attachments in email
                        foreach (var files in db.TRF_File.Where(o => o.TravelID == TID && o.UserID == UID))
                        {
                            byte[] attachemnt = files.Content;
                            string filename = files.FileName;

                            if (attachemnt != null)
                            {
                                mm.Attachments.Add(new Attachment(new System.IO.MemoryStream(attachemnt), filename));
                            }
                        }
                        //converting to pdf and adding as attachment
                        var Renderer = new IronPdf.HtmlToPdf();
                        var PDF = Renderer.RenderHTMLFileAsPdf(Server.MapPath("~/EmailBody.html"));
                        PDF.PrependPdf(Renderer.RenderHtmlAsPdf(CreateBody(UID, TID, UTID)));
                        PDF.RemovePage(PDF.PageCount - 1);
                        PDF.SaveAs(OutputPath);
                        mm.Attachments.Add(new Attachment(OutputPath));
                        //smtp settings
                        smtp.Host = "mailrelay.westminster-mo.edu";
                        smtp.EnableSsl = false;
                        smtp.Port = 25;
                        smtp.Send(mm);

                    }
                    System.IO.File.Delete(OutputPath);
                }

            }

        }

        //creating the body for email
        private string CreateBody(Guid UID, Guid TID, Guid UTID)
        {
            string body = string.Empty;
            using (System.IO.StreamReader reader = new System.IO.StreamReader(Server.MapPath("~/EmailBody.html")))
            {
                body = reader.ReadToEnd();

            }

            if (db.TRF_UserTravelExpense.Any(o => o.UserID == UID && o.TravelID == TID && o.ID == UTID))
            {

                //Requestor's information
                TRF_User users = db.TRF_User.Find(UID);
                if (users.TravelID == TID)
                {
                    body = body.Replace("{name}", users.FirstName + " " + users.LastName);
                    body = body.Replace("{type}", users.UserType);
                    body = body.Replace("{email}", users.Email);
                    body = body.Replace("{department}", users.Department);
                    body = body.Replace("{chair}", users.SupervisorName);
                    body = body.Replace("{Account}", users.AccountCode);
                    body = body.Replace("{notescomments}", users.notescomments);
                }
                //Requestor's information
                TRF_TravelAddress usertravel = db.TRF_TravelAddress.Find(TID);
                body = body.Replace("{departuredate}", usertravel.DepartureDate.ToShortDateString());
                body = body.Replace("{arrivaldate}", usertravel.ArrivalDate.ToShortDateString());
                body = body.Replace(" {address}", usertravel.Address + ", " + usertravel.City + " " + usertravel.State + ", " + usertravel.ZipCode);

                //calculating the travel cost
                double lodging = 0.00, autofare = 0.00, airfare = 0.00, meal = 0.00, regestration = 0.00, misci = 0.00, total = 0.00;
                foreach (var expenses in db.TRF_UserTravelExpense.Where(o => o.ID == UTID))
                {
                    if (expenses.ExpenseID == new Guid("74C4127D-59E0-4751-903E-4EF765AAC395"))
                    {
                        regestration = expenses.Amount.Value;
                    }
                    if (expenses.ExpenseID == new Guid("758A5FA1-0F7D-4173-BF30-798A82D7FA77"))
                    {
                        lodging = expenses.Amount.Value;
                    }
                    if (expenses.ExpenseID == new Guid("B35FB5EE-9459-4362-B968-8BB4325E3D25"))
                    {
                        misci = expenses.Amount.Value;
                    }
                    if (expenses.ExpenseID == new Guid("A9F6D7B2-367D-4CF6-B806-A5EB146E4269"))
                    {
                        autofare = expenses.Amount.Value;
                    }
                    if (expenses.ExpenseID == new Guid("D0D30F46-8670-45CB-B4A8-D499FDD1FFC3"))
                    {
                        meal = expenses.Amount.Value;
                    }
                    if (expenses.ExpenseID == new Guid("88E6107B-EE04-4B3B-B226-E01EFFF1B202"))
                    {
                        airfare = expenses.Amount.Value;
                    }
                }

                total = lodging + autofare + airfare + meal + regestration + misci;
                //Travel Expenses
                body = body.Replace("{lodging}", "$" + String.Format("{0:0.00#}", lodging));
                body = body.Replace("{autofate}", "$" + String.Format("{0:0.00#}", autofare));
                body = body.Replace("{airfare}", "$" + String.Format("{0:0.00#}", airfare));
                body = body.Replace("{mealcost}", "$" + String.Format("{0:0.00#}", meal));
                body = body.Replace("{registcost}", "$" + String.Format("{0:0.00#}", regestration));
                body = body.Replace("{Misce}", "$" + String.Format("{0:0.00#}", misci));
                body = body.Replace("{total}", "$" + String.Format("{0:0.00#}", total));
            }

            return body;
        }


      
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
