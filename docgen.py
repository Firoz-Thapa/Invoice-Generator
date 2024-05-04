from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")

invoice_list = [[2, "pen", 0.5, 1],
                [1, "paper pack", 5, 5],
                [2, "notebook", 2, 4]]

doc.render({"name":"gyawat",
            "phone": "+954-0202033",
            "invoice_list": invoice_list,
            "subtotal": 10,
            "salestax":"10%",
            "total":9})
doc.save("new_invoice.docx")