# CMP_CB1_MAGENTA

OTEREV.CMP_CB1_MAGENTA

1. Digital_Order_Id - Μοναδικό αναγνωριστικό παραγγελίας ή αιτήματος
2. Document_file_info_id - Μοναδικό αναγνωριστικό του CMP
3. Topic - Καρφωτή τιμή που χρησιμοποιείται για να διαχωρίζεται η ροή από την οποία προκύπτουν τα έγγραφα
4. Category - Καρφωτή τιμή που υποδεικνύει την κατηγορία του εγγράφου
5. Subcategory - Καρφωτή τιμή που υποδεικνύει την υπό-κατηγορία του εγγράφου
6. Netapp - Always 0 όταν γίνεται insert στον πίνακα. Ενημερώνεται σε 1 εάν το import στον NETAPP είναι επιτυχές.
7. File_path - FilePath + File_Name_with_No_Extension π.χ.: /folder1/folder2/aitisi12345
8. File_type - π.χ.: .pdf,. jpg, .tif
9. Customer_Code - π.χ.: OTEC984432
10. Billing_account_id - π.χ.: 1-2BNU97QI (αριθμός λογαριασμού πελάτη)
11. Shop_code - π.χ.: 10421.00001 - COSMOTE ΓΡΕΒΕΝΩΝ
12. Doc_Date - Ημερομηνία εισαγωγής στον DB-Link πίνακα
13. Update_Date - Ημερομηνία όπου γίνεται το update του progress της εισαγωγής
14. Import_Status - Always 0 όταν γίνεται insert στον πίνακα. Ενημερώνεται σε 1 (σε δεύτερο χρόνο από την FileNet) εάν το import είναι επιτυχές.
15. F_Docnumber - Always null όταν γίνεται insert στον πίνακα. Ενημερώνεται σε δεύτερο χρόνο από την FileNet.

Το path που θα μεταφέρονται τα αρχεία θα είναι \\10.101.6.9\vol_edocs3_prod\CB1_MAGENTA