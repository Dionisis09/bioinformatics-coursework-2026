from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

BASE_DIR = Path(r"C:\Users\dioni\Desktop\general\viopliroforiki ergasia 1")
REPORT_DIR = BASE_DIR / "report"
SCREENSHOTS_DIR = Path(r"C:\Users\dioni\Pictures\vioplitoforiki screenshots")
OUTPUT_PATH = REPORT_DIR / "viopliroforiki_parousiasi_ellinika_enriched_v2.docx"

SECTIONS = [
    {
        "title": "Βιοπληροφορική - Εργασία 1",
        "intro": [
            "Παρουσίαση της μεθοδολογίας, της εκτέλεσης του κώδικα και των βασικών αποτελεσμάτων της εργασίας.",
            "Σε κάθε ενότητα παρουσιάζονται οι εντολές που τρέξαμε, τα σχετικά screenshots και το βιολογικό συμπέρασμα που προκύπτει.",
        ],
    },
    {
        "title": "Εισαγωγή και στόχοι",
        "paragraphs": [
            "Η εργασία συνδυάζει πρακτικά εργαλεία βιοπληροφορικής με θεωρητικές έννοιες. Ο βασικός στόχος ήταν να μελετήσουμε αναζητήσεις ομόλογων πρωτεϊνών, στοιχίσεις αλληλουχιών, σύγκριση δομών και μεθόδους indexing.",
            "Παράλληλα, η εργασία μας έδωσε την ευκαιρία να δούμε πώς συνδέεται η υπολογιστική ανάλυση με πραγματικά βιολογικά ερωτήματα, όπως η διατήρηση μιας πρωτεΐνης, η ταυτοποίηση μιας άγνωστης αλληλουχίας και η σύγκριση ιικών πρωτεϊνών.",
        ],
        "bullets": [
            "Αναζήτηση ομόλογων πρωτεϊνών με BLASTP",
            "Multiple sequence alignment με MAFFT, MUSCLE και T-Coffee",
            "Pairwise και structural alignment για spike πρωτεΐνες",
            "Υλοποίηση global και local alignment με δυναμικό προγραμματισμό",
            "Θεωρητική προσέγγιση suffix tree και τεχνικών indexing",
        ],
        "notes": [
            "Εδώ μπορώ να τονίσω ότι η εργασία δεν είναι απλώς μια σειρά από commands, αλλά μια αλυσίδα ερωτημάτων όπου κάθε βήμα παράγει δεδομένα για το επόμενο.",
            "Επίσης αξίζει να πω ότι το ενδιαφέρον δεν είναι μόνο τεχνικό, αλλά και βιολογικό: θέλουμε να καταλάβουμε τι σημαίνει η ομοιότητα ή η διαφορά ανάμεσα στις αλληλουχίες και στις δομές.",
        ],
    },
    {
        "title": "Ερώτημα 1(i): AlphaFold 3 και ESMFold",
        "paragraphs": [
            "Στο πρώτο υποερώτημα έγινε σύντομη θεωρητική σύγκριση δύο σύγχρονων εργαλείων πρόβλεψης πρωτεϊνικής δομής: του AlphaFold 3 και του ESMFold.",
            "Το AlphaFold 3 είναι πιο γενικό και επεκτείνει την πρόβλεψη σε ευρύτερα βιομοριακά συστήματα, ενώ το ESMFold είναι ιδιαίτερα γρήγορο και χρήσιμο όταν θέλουμε άμεση πρόβλεψη δομής από μια πρωτεϊνική αλληλουχία.",
            "Το βασικό συμπέρασμα είναι ότι και τα δύο εργαλεία είναι πολύ σημαντικά, αλλά εξυπηρετούν ελαφρώς διαφορετικές ανάγκες ανάλογα με το πρόβλημα που θέλουμε να μελετήσουμε.",
        ],
        "notes": [
            "Αυτό δείχνει ότι έχω καταλάβει πως στη βιοπληροφορική δεν υπάρχει ένα μόνο 'καλύτερο' εργαλείο. Η επιλογή εξαρτάται από το αν προέχει η ταχύτητα, η γενικότητα ή ο τύπος του βιομορίου που μελετάμε.",
            "Μπορώ επίσης να σχολιάσω ότι η πρόβλεψη δομής είναι σήμερα κρίσιμη, γιατί δίνει βιολογικές πληροφορίες ακόμα και όταν δεν έχουμε πειραματικά δεδομένα.",
        ],
    },
    {
        "title": "Ερώτημα 1(ii): BLASTP για hemoglobin beta",
        "commands": [r"python src\q1_blastp_collect.py"],
        "images": [
            ("q1 ii results.png", "Εκτέλεση του script για BLASTP της ανθρώπινης hemoglobin beta."),
            ("q1 ii results top10 file.png", r"Output file: results\question1\hemoglobin_beta_top10.fasta"),
        ],
        "paragraphs": [
            "Σε αυτό το βήμα χρησιμοποιήσαμε ως query την ανθρώπινη hemoglobin subunit beta με accession NP_000509 και εκτελέσαμε BLASTP για να εντοπίσουμε σχετικές πρωτεϊνικές αλληλουχίες.",
            "Το script επέλεξε 10 κατάλληλα hits και τα αποθήκευσε σε αρχείο FASTA. Με αυτόν τον τρόπο δημιουργήσαμε ένα καθαρό dataset εισόδου για το επόμενο βήμα του multiple sequence alignment.",
            "Τα hits προέρχονται από πολύ συγγενικά είδη, κάτι που ήδη προϊδεάζει ότι η πρωτεΐνη αυτή είναι έντονα συντηρημένη εξελικτικά.",
        ],
        "bullets": [
            "Query accession: NP_000509",
            "Επιλογή 10 σχετικών πρωτεϊνικών hits",
            "Αποθήκευση σε FASTA για περαιτέρω ανάλυση",
        ],
        "notes": [
            "Εδώ φαίνεται ότι το BLASTP δεν χρησιμοποιείται απλώς για να βρούμε 'παρόμοια' sequences, αλλά για να χτίσουμε ένα λογικό σύνολο δεδομένων που θα αναλυθεί εξελικτικά στο επόμενο βήμα.",
            "Το γεγονός ότι τα πρώτα hits είναι από συγγενικά πρωτεύοντα είναι επίσης βιολογικά αναμενόμενο, άρα το αποτέλεσμα έχει και βιολογική συνέπεια, όχι μόνο τεχνική επιτυχία.",
        ],
    },
    {
        "title": "Ερώτημα 1(iii): Multiple Sequence Alignment",
        "commands": [
            r"python tools\ebi_clients\mafft.py --email up1093449@ac.upatras.gr --stype protein --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\mafft_test",
            r"python tools\ebi_clients\muscle.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\muscle_run",
            r"python tools\ebi_clients\tcoffee.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\tcoffee_run",
        ],
        "images": [
            ("q1(iii) 1 terminal.png", "Εκτέλεση των jobs για MAFFT, MUSCLE και T-Coffee."),
            ("q1(iii) 2 result file.png", r"Alignment output: results\question1\msa\muscle_run.aln-clustalw.clw"),
        ],
        "paragraphs": [
            "Αφού συγκεντρώσαμε τις 10 αλληλουχίες, πραγματοποιήσαμε multiple sequence alignment με τρία διαφορετικά εργαλεία: MAFFT, MUSCLE και T-Coffee.",
            "Στο terminal φαίνεται ότι και τα τρία jobs ολοκληρώθηκαν επιτυχώς και παρήγαγαν τα αντίστοιχα αρχεία αποτελεσμάτων. Στο alignment του MUSCLE βλέπουμε καθαρά πολλές πλήρως συντηρημένες θέσεις, κάτι που υποδηλώνεται και από τις συνεχόμενες στήλες με αστερίσκους.",
            "Για να καλυφθεί και η απαίτηση της εκφώνησης για δοκιμή παραμέτρων, έγινε επιπλέον ένα δεύτερο MAFFT run με την επιλογή maxiterate=2. Το εναλλακτικό αποτέλεσμα αποθηκεύτηκε στα αρχεία results/question1/msa/mafft_test_maxiterate2.aln-fasta.fasta και results/question1/msa/mafft_test_maxiterate2.submission.params.",
            "Η σύγκριση του αρχικού MAFFT alignment με το εναλλακτικό run έδειξε ότι για αυτό το dataset το αποτέλεσμα παρέμεινε ίδιο. Αυτό είναι αναμενόμενο, επειδή οι αλληλουχίες είναι ήδη πολύ όμοιες και η στοίχιση είναι αρκετά σταθερή.",
            "Το συμπέρασμα του ερωτήματος είναι ότι η hemoglobin beta εμφανίζει πολύ υψηλή ομοιότητα στα επιλεγμένα είδη, άρα πρόκειται για ιδιαίτερα συντηρημένη πρωτεΐνη.",
        ],
        "bullets": [
            "Χρήση τριών διαφορετικών MSA εργαλείων",
            "Παραγωγή alignment και συνοδευτικών αρχείων",
            "Έλεγχος εναλλακτικής παραμέτρου MAFFT (maxiterate=2)",
            "Ισχυρή εξελικτική συντήρηση της πρωτεΐνης",
        ],
        "notes": [
            "Στην παρουσίαση μπορώ να επισημάνω ότι χρησιμοποίησα τρία διαφορετικά εργαλεία όχι επειδή περιμένω ριζικά διαφορετικό αποτέλεσμα, αλλά για να δείξω ότι το μοτίβο συντήρησης είναι σταθερό ανεξάρτητα από τη μέθοδο.",
            "Επιπλέον, το δοκιμαστικό run με διαφορετική παράμετρο στο MAFFT δείχνει ότι δεν ακολουθήσαμε μηχανικά τα defaults, αλλά ελέγξαμε αν η στοίχιση αλλάζει όταν τροποποιείται η διαδικασία refinement.",
            "Οι αστερίσκοι στο alignment σημαίνουν πλήρη συντήρηση θέσεων και αυτό είναι ένδειξη ότι συγκεκριμένα κατάλοιπα είναι σημαντικά για τη λειτουργία της πρωτεΐνης.",
        ],
    },
    {
        "title": "Ερώτημα 2: Cytochrome c στα Mammalia",
        "source": "Source: UniProt entry P99999 (CYC_HUMAN)",
        "commands": [r"python src\q2_cytochrome_c.py"],
        "images": [
            ("q2 3.png", "Σελίδα UniProt για την ανθρώπινη Cytochrome c."),
            ("q2 1.png", "Εκτέλεση του script για BLASTP στα Mammalia."),
            ("q2 2.png", r"Output file: results\question2\summary.txt"),
        ],
        "paragraphs": [
            "Στο δεύτερο ερώτημα ξεκινήσαμε από την ανθρώπινη Cytochrome c, χρησιμοποιώντας το UniProt accession P99999 ως αφετηρία για την ανάλυση.",
            "Έπειτα εκτελέσαμε BLASTP με περιορισμό στα θηλαστικά και υπολογίσαμε το ποσοστό ταυτότητας κάθε αλληλουχίας ως προς την ανθρώπινη. Στο summary φαίνεται ότι τα ποσοστά identity είναι εξαιρετικά υψηλά.",
            "Το βιολογικό συμπέρασμα είναι ότι η Cytochrome c είναι πολύ έντονα συντηρημένη στα Mammalia, γεγονός που είναι αναμενόμενο επειδή συμμετέχει σε θεμελιώδεις κυτταρικές λειτουργίες.",
        ],
        "bullets": [
            "Αρχική πηγή: UniProt P99999",
            "BLASTP με περιορισμό στα Mammalia",
            "Identity περίπου 98% έως 100%",
        ],
        "notes": [
            "Εδώ είναι καλό να πω ότι η Cytochrome c χρησιμοποιείται κλασικά ως παράδειγμα εξελικτικά συντηρημένης πρωτεΐνης, άρα τα υψηλά percentages δεν είναι τυχαία.",
            "Μπορώ επίσης να σχολιάσω ότι η υψηλή συντήρηση υποδηλώνει ισχυρούς λειτουργικούς περιορισμούς: μεγάλες μεταβολές πιθανότατα δεν είναι εύκολα ανεκτές από το κύτταρο.",
        ],
    },
    {
        "title": "Ερώτημα 3: Ταυτοποίηση άγνωστης αλληλουχίας",
        "commands": [r"python src\q3_identify_protein.py"],
        "images": [
            ("q3 1.png", "Εκτέλεση του script για BLASTP της άγνωστης αλληλουχίας."),
            ("q3 2.png", "Συνέχεια των top BLAST hits στο terminal."),
            ("q3 3 summary.png", r"Summary file: results\question3\summary.txt (μέρος 1)"),
            ("q3 4 summary .png", r"Summary file: results\question3\summary.txt (μέρος 2)"),
        ],
        "paragraphs": [
            "Στο τρίτο ερώτημα δόθηκε η σύντομη αλληλουχία GIVEQCCTSICSLYQLENYCN και ο στόχος ήταν να προσδιορίσουμε σε ποια πρωτεΐνη αντιστοιχεί.",
            "Η αναζήτηση με BLASTP έδειξε πολύ καθαρά ότι η αλληλουχία αντιστοιχεί σε Insulin A chain. Το πρώτο hit προέρχεται από Sus scrofa, ενώ στα κορυφαία αποτελέσματα υπάρχουν και ακριβή ανθρώπινα hits.",
            "Άρα η πρωτεΐνη ταυτοποιείται με ασφάλεια ως insulin A chain, όμως η αλληλουχία είναι τόσο μικρή και τόσο συντηρημένη ώστε δεν μπορεί να θεωρηθεί μοναδικός δείκτης ενός μόνο οργανισμού.",
            "Η γνωστή λειτουργία της ινσουλίνης είναι η ρύθμιση της γλυκόζης στο αίμα, άρα πρόκειται για πρωτεΐνη με πολύ σημαντικό φυσιολογικό ρόλο.",
        ],
        "bullets": [
            "Protein identification: Insulin A chain",
            "Best first hit: Sus scrofa",
            "Υπάρχουν και ακριβή ανθρώπινα hits στα κορυφαία αποτελέσματα",
        ],
        "notes": [
            "Αυτό είναι καλό σημείο για να δείξω κριτική σκέψη: δεν αρκεί να πω απλώς 'βρέθηκε ινσουλίνη', αλλά πρέπει να εξηγήσω γιατί το πρώτο hit δεν σημαίνει απαραίτητα μοναδική προέλευση από έναν οργανισμό.",
            "Το μικρό μήκος της αλληλουχίας είναι καθοριστικό. Όταν ένα peptide είναι τόσο σύντομο και πολύ συντηρημένο, η ταυτοποίηση της πρωτεΐνης είναι ασφαλής, αλλά η ταυτοποίηση του είδους είναι πιο αδύναμη.",
        ],
    },
    {
        "title": "Ερώτημα 4(i): Pairwise alignment spike πρωτεϊνών",
        "commands": [
            r"python src\q4_spike_alignment.py",
            r"python tools\ebi_clients\emboss_needle.py --email up1093449@ac.upatras.gr --stype protein --asequence results\question4\sars_cov_2_spike.fasta --bsequence results\question4\sars_cov_spike.fasta --outfile results\question4\needle_run",
        ],
        "images": [
            ("q4 i 1 .png", "Προετοιμασία των αλληλουχιών και εκτέλεση του EMBOSS Needle."),
            ("q4 i txt 2.png", r"Output file: results\question4\needle_run.aln.txt (summary metrics)"),
            ("q4 i txt 3.png", "Αναλυτικό pairwise alignment από το EMBOSS Needle."),
        ],
        "paragraphs": [
            "Στο πρώτο μέρος του Ερωτήματος 4 συγκρίναμε τις spike πρωτεΐνες του SARS-CoV-2 και του SARS-CoV. Αρχικά έγινε η ανάκτηση των αλληλουχιών και στη συνέχεια εκτελέστηκε pairwise alignment με το EMBOSS Needle.",
            "Στο summary του alignment παρατηρούμε identity 76.4%, similarity 87.0% και gaps 2.0%. Αυτές οι τιμές δείχνουν ότι οι δύο αλληλουχίες είναι σαφώς συγγενικές, αλλά δεν είναι ίδιες.",
            "Το detailed alignment επιτρέπει να δούμε σε ποια τμήματα υπάρχει ισχυρή ομοιότητα και σε ποια σημεία εμφανίζονται διαφορές. Αυτό είναι ιδιαίτερα σημαντικό, επειδή μικρές αλλαγές στη spike πρωτεΐνη μπορούν να επηρεάσουν την αλληλεπίδραση με τον ξενιστή.",
        ],
        "bullets": [
            "Pairwise alignment με EMBOSS Needle",
            "Identity: 76.4%",
            "Similarity: 87.0%",
            "Βιολογικό συμπέρασμα: υψηλή συγγένεια αλλά όχι ταύτιση",
        ],
        "notes": [
            "Εδώ μπορώ να εξηγήσω ότι το identity μετρά ακριβείς ίδιους χαρακτήρες, ενώ το similarity επιτρέπει και βιοχημικά συγγενικές αντικαταστάσεις. Γι' αυτό το similarity είναι υψηλότερο από το identity.",
            "Η διαφορά αυτή είναι ουσιαστική, γιατί δείχνει ότι οι δύο spike πρωτεΐνες διατηρούν παρόμοιο λειτουργικό χαρακτήρα, παρότι έχουν αρκετές μεταλλάξεις σε επίπεδο sequence.",
        ],
    },
    {
        "title": "Ερώτημα 4(ii): Δομική σύγκριση spike πρωτεϊνών",
        "commands": [r"python src\q4_structure_alignment.py"],
        "images": [
            (r"results\question4\swissmodel_p0dtc2.png", "SWISS-MODEL καταχώριση για το SARS-CoV-2 spike (P0DTC2)."),
            (r"results\question4\swissmodel_p59594.png", "SWISS-MODEL καταχώριση για το SARS-CoV spike (P59594)."),
            ("q4 ii 1 terminal.png", "Εκτέλεση του structure alignment script."),
            ("q4 ii 2 png.png", r"Output image: results\question4\rcsb_alignment_summary.png"),
        ],
        "paragraphs": [
            "Στο δεύτερο μέρος του Ερωτήματος 4 δεν περιοριστήκαμε στην αλληλουχία, αλλά συγκρίναμε και τις τρισδιάστατες δομές των spike πρωτεϊνών με χρήση του RCSB structure alignment backend.",
            "Για να συνδεθεί πλήρως η παρουσίαση με την εκφώνηση, συμπεριλάβαμε και τα δύο οπτικά τεκμήρια από το SWISS-MODEL Repository, ένα για το μοντέλο του SARS-CoV-2 spike και ένα για το μοντέλο του SARS-CoV spike. Έτσι φαίνεται καθαρά ότι εξετάστηκαν και οι δύο δομές πριν από τη δομική ευθυγράμμιση.",
            "Τα κύρια metrics ήταν TM-score 0.95, RMSD 1.8, sequence identity 0.75 και πολύ υψηλό coverage. Αυτοί οι δείκτες δείχνουν ότι οι δύο δομές είναι πολύ κοντινές στον τρισδιάστατο χώρο.",
            "Άρα, παρόλο που σε επίπεδο αλληλουχίας υπάρχουν διαφορές, σε επίπεδο συνολικής αναδίπλωσης και δομικής οργάνωσης οι spike πρωτεΐνες παραμένουν εξαιρετικά όμοιες.",
        ],
        "bullets": [
            "Οπτική επιβεβαίωση και των δύο μοντέλων από το SWISS-MODEL",
            "TM-score: 0.95",
            "RMSD: 1.8",
            "Coverage: 94% / 97%",
            "Συμπέρασμα: οι δομές είναι πολύ κοντινές σε 3D επίπεδο",
        ],
        "notes": [
            "Αυτό είναι πολύ δυνατό σημείο στην παρουσίαση, γιατί δείχνει ότι sequence similarity και structural similarity δεν είναι ταυτόσημες έννοιες. Μπορεί να υπάρχουν διαφορές στην αλληλουχία, αλλά η συνολική δομή να παραμένει πολύ σταθερή.",
            "Το ότι παρουσιάζονται και τα δύο SWISS-MODEL screenshots βοηθά να φανεί ότι δεν βασιστήκαμε μόνο σε ένα τελικό metric, αλλά ελέγξαμε τα ίδια τα μοντέλα που συγκρίνονται.",
            "Το υψηλό TM-score σημαίνει ότι οι πρωτεΐνες έχουν πολύ κοντινή παγκόσμια αναδίπλωση, ενώ το RMSD δείχνει ότι οι αποκλίσεις σε τρισδιάστατο χώρο είναι σχετικά μικρές.",
        ],
    },
    {
        "title": "Ερώτημα 5: Global και Local alignment με δυναμικό προγραμματισμό",
        "commands": [r"python src\q5_alignment_dp.py"],
        "images": [
            ("q5 1.png", "Εκτέλεση του script για Needleman-Wunsch και Smith-Waterman."),
            ("q5 2 txt global.png", r"Output file: results\question5\global_alignment.txt"),
            ("q5 3 txt local.png", r"Output file: results\question5\local_alignment.txt"),
        ],
        "paragraphs": [
            "Στο πέμπτο ερώτημα υλοποιήσαμε δύο κλασικούς αλγορίθμους δυναμικού προγραμματισμού για στοίχιση ακολουθιών: τον Needleman-Wunsch για global alignment και τον Smith-Waterman για local alignment.",
            "Στο global alignment το βέλτιστο score ήταν -2, κάτι που δείχνει ότι όταν προσπαθούμε να στοιχίσουμε ολόκληρες τις δύο αλληλουχίες, οι διαφορές τους επιβαρύνουν το συνολικό αποτέλεσμα.",
            "Στο local alignment το βέλτιστο score ήταν 3 και το καλύτερο κοινό τοπικό μοτίβο ήταν το GTG. Αυτό δείχνει πολύ καθαρά τη διαφορά ανάμεσα στη συνολική και την τοπική στοίχιση: η local αναζητά το καλύτερο κοινό υποτμήμα, ακόμα κι αν οι συνολικές αλληλουχίες διαφέρουν αρκετά.",
        ],
        "bullets": [
            "Global alignment score: -2",
            "Local alignment score: 3",
            "Καλύτερο τοπικό motif: GTG",
        ],
        "notes": [
            "Εδώ μπορώ να δείξω ότι έχω καταλάβει ουσιαστικά τη διαφορά των δύο αλγορίθμων: η global alignment προσπαθεί να ευθυγραμμίσει όλο το μήκος, ενώ η local ψάχνει μόνο το καλύτερο κοινό τμήμα.",
            "Αυτό εξηγεί γιατί η local alignment είναι συχνά πιο χρήσιμη όταν δύο αλληλουχίες μοιράζονται μόνο ένα κοινό functional motif και όχι συνολική ομοιότητα σε όλο το μήκος τους.",
        ],
    },
    {
        "title": "Ερώτημα 6(α): Θεωρητική λύση με suffix tree",
        "source": r"Theory file: report\question6_theory.md",
        "images": [
            ("q6 1.png", "Θεωρητική λύση με suffix tree και αποθήκευση minPos."),
        ],
        "paragraphs": [
            "Στο πρώτο μέρος του Ερωτήματος 6 ζητήθηκε θεωρητική λύση για αποδοτική απάντηση ερωτημάτων εμφάνισης προτύπου πριν από μια θέση k.",
            "Η πρόταση που δόθηκε βασίζεται σε suffix tree. Η βασική ιδέα είναι ότι για κάθε κόμβο αποθηκεύεται η μικρότερη θέση εμφάνισης minPos ανάμεσα στα φύλλα του υποδέντρου του.",
            "Με αυτόν τον τρόπο πετυχαίνουμε preprocessing O(n) και, για ένα query μήκους m, χρόνο απάντησης O(m). Αυτό καθιστά τη μέθοδο θεωρητικά πολύ αποδοτική για επαναλαμβανόμενα queries πάνω στο ίδιο κείμενο.",
        ],
        "bullets": [
            "Suffix tree του T$",
            "Αποθήκευση minPos σε κάθε κόμβο",
            "Preprocessing O(n)",
            "Query O(m)",
        ],
        "notes": [
            "Σε αυτό το σημείο μπορώ να δείξω ότι δεν απομνημόνευσα απλώς μια δομή δεδομένων, αλλά κατάλαβα γιατί δουλεύει: η πληροφορία minPos συμπυκνώνει το αν υπάρχει εμφάνιση πριν από ένα όριο k χωρίς να χρειάζεται να ξαναψάχνω όλο το κείμενο.",
            "Άρα η προεπεξεργασία επενδύει χρόνο μία φορά και μετά τα επαναλαμβανόμενα queries απαντώνται πολύ πιο γρήγορα.",
        ],
    },
    {
        "title": "Ερώτημα 6(β): Bloom filters, Suffix Arrays και Burrows-Wheeler",
        "source": r"Theory file: report\question6_theory.md",
        "images": [
            ("q6 2.png", "Σύγκριση Bloom filters, suffix arrays και Burrows-Wheeler based indexing."),
        ],
        "paragraphs": [
            "Στο δεύτερο μέρος έγινε σύντομη διερεύνηση τριών σημαντικών οικογενειών τεχνικών indexing: Bloom filters, suffix arrays και Burrows-Wheeler based indexing.",
            "Τα Bloom filters είναι πολύ ελαφριά και ιδανικά για membership queries, αλλά δεν δίνουν πλήρες και ακριβές indexing. Τα suffix arrays προσφέρουν ακριβές full-text indexing, αλλά απαιτούν περισσότερο χώρο.",
            "Οι BWT και FM-index προσεγγίσεις προσπαθούν να συνδυάσουν ακριβές indexing με πολύ καλύτερη συμπίεση. Αυτός είναι και ο λόγος που τέτοιες τεχνικές είναι τόσο σημαντικές για μεγάλα βιολογικά datasets, όπου η οικονομία χώρου παίζει κρίσιμο ρόλο.",
        ],
        "bullets": [
            "Bloom filters: μικρός χώρος, membership queries",
            "Suffix arrays: ακριβές indexing",
            "BWT / FM-index: ακριβές indexing με καλύτερη συμπίεση",
        ],
        "notes": [
            "Εδώ μπορώ να πω ότι οι τρεις τεχνικές δεν είναι ανταγωνιστικές με τον ίδιο τρόπο, αλλά καλύπτουν διαφορετικούς συμβιβασμούς ανάμεσα σε χώρο, ταχύτητα και ακρίβεια.",
            "Αυτό είναι ιδιαίτερα χρήσιμο στη βιοπληροφορική, γιατί τα datasets είναι τεράστια και συχνά η επιλογή δομής indexing είναι εξίσου σημαντική με τον ίδιο τον αλγόριθμο αναζήτησης.",
        ],
    },
    {
        "title": "Συνολικά συμπεράσματα της εργασίας",
        "paragraphs": [
            "Η εργασία κάλυψε ένα ευρύ φάσμα θεμάτων βιοπληροφορικής, από την ανάλυση αλληλουχιών μέχρι τη σύγκριση δομών και τις θεωρητικές δομές indexing.",
            "Στην πράξη χρησιμοποιήσαμε BLASTP, multiple alignment εργαλεία, EMBOSS Needle, RCSB structure alignment και υλοποιήσεις δυναμικού προγραμματισμού. Παράλληλα, μελετήσαμε και θεωρητικές τεχνικές όπως το suffix tree και οι BWT-based index structures.",
            "Το γενικό συμπέρασμα είναι ότι η βιοπληροφορική απαιτεί τόσο σωστή βιολογική ερμηνεία όσο και σωστή υπολογιστική μεθοδολογία. Τα αποτελέσματα της εργασίας ήταν συνεπή και έδειξαν στην πράξη πώς συνδέονται αυτά τα δύο επίπεδα.",
        ],
        "bullets": [
            "Σύνδεση sequence analysis, structure comparison και indexing",
            "Συνδυασμός πρακτικών εργαλείων και θεωρητικών αλγορίθμων",
            "Βιολογικά και υπολογιστικά συνεπή αποτελέσματα",
        ],
        "notes": [
            "Κλείνοντας, μπορώ να τονίσω ότι η εργασία δείχνει μια πλήρη λογική ροή: από την αναζήτηση ομόλογων, στην ποσοτική σύγκριση αλληλουχιών, στη σύγκριση δομών και τελικά στη θεωρητική οργάνωση μεγάλων βιολογικών δεδομένων.",
            "Αυτό βοηθά τον καθηγητή να δει ότι δεν αντιμετώπισα τα ερωτήματα αποσπασματικά, αλλά ως ενιαία εφαρμογή εργαλείων και ιδεών της βιοπληροφορικής.",
        ],
    },
]


def set_normal_style(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)


def add_intro(document: Document) -> None:
    title_paragraph = document.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_paragraph.add_run(SECTIONS[0]["title"])
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(22)

    for line in SECTIONS[0]["intro"]:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.add_run(line)

    document.add_page_break()


def add_commands(document: Document, commands: list[str]) -> None:
    paragraph = document.add_paragraph()
    paragraph.add_run("Εντολή που τρέχουμε:").bold = True

    for command in commands:
        command_paragraph = document.add_paragraph()
        run = command_paragraph.add_run(command)
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)


def add_images(document: Document, images: list[tuple[str, str]]) -> None:
    for image_name, caption in images:
        image_path = Path(image_name)
        if not image_path.is_absolute():
            screenshot_candidate = SCREENSHOTS_DIR / image_name
            project_candidate = BASE_DIR / image_name
            if screenshot_candidate.exists():
                image_path = screenshot_candidate
            else:
                image_path = project_candidate
        caption_paragraph = document.add_paragraph()
        caption_run = caption_paragraph.add_run(caption)
        caption_run.italic = True

        if image_path.exists():
            document.add_picture(str(image_path), width=Inches(6.0))
        else:
            missing = document.add_paragraph()
            missing.add_run(f"Λείπει screenshot: {image_name}").bold = True



def add_paragraphs(document: Document, paragraphs: list[str]) -> None:
    for text in paragraphs:
        document.add_paragraph(text)



def add_bullets(document: Document, bullets: list[str]) -> None:
    for bullet in bullets:
        document.add_paragraph(bullet, style="List Bullet")


def add_notes(document: Document, notes: list[str]) -> None:
    heading = document.add_paragraph()
    heading.add_run("Σχόλια παρουσίασης / τι δείχνει ότι έχω καταλάβει:").bold = True
    for note in notes:
        document.add_paragraph(note, style="List Bullet")



def build_document() -> Path:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    set_normal_style(document)
    add_intro(document)

    content_sections = SECTIONS[1:]
    for index, content in enumerate(content_sections):
        document.add_heading(content["title"], level=1)

        source = content.get("source")
        if source:
            source_paragraph = document.add_paragraph()
            source_run = source_paragraph.add_run(source)
            source_run.italic = True

        commands = content.get("commands", [])
        if commands:
            add_commands(document, commands)

        images = content.get("images", [])
        if images:
            add_images(document, images)

        paragraphs = content.get("paragraphs", [])
        if paragraphs:
            add_paragraphs(document, paragraphs)

        bullets = content.get("bullets", [])
        if bullets:
            add_bullets(document, bullets)

        notes = content.get("notes", [])
        if notes:
            add_notes(document, notes)

        if index != len(content_sections) - 1:
            document.add_page_break()

    document.save(OUTPUT_PATH)
    return OUTPUT_PATH


if __name__ == "__main__":
    path = build_document()
    print(path)
