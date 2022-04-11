import pandas as pd
import os
import sys


remit_order = ["Bellboat",
               "Very sheltered water – Kayak",
               "Very sheltered water – Canoe",
               "Very sheltered water – SUP",
               "Sheltered water – Kayak",
               "Sheltered water – Canoe",
               "Sheltered water – SUP",
               "Moderate white water – Kayak",
               "Moderate white water – Canoe",
               "Moderate white water – SUP",
               "Moderate open water – Kayak",
               "Moderate open water – Canoe",
               "Moderate open water – SUP",
               "Moderate sea – kayak",
               "Moderate sea – canoe",
               "Moderate sea – SUP",
               "Advanced white water – kayak",
               "Advanced white water – canoe",
               "Advanced open water – canoe",
               "Advanced sea – kayak"]
               
               
provider_options = ["PPA Provider", 
                    "Paddle Explore Award Provider",
                    "White Water Award Provider",
                    "Canoe Award Provider",
                    "Touring Award Provider",
                    "Open Water Touring Award Provider",
                    "Multi Day Touring Award Provider",
                    "FSRT Provider"]


def non_nan(stuff):
    return [a for a in stuff if str(a) not in ["nan", "NaT"]]


def non_nan_date(stuff):
    values = []
    for a in stuff: 
        if str(a) not in ["nan", "NaT"]:
            try:
                values.append(a.date())
            except AttributeError:
                values.append(a)
    return values


def add_to_writer_with_fixed_col_widths(writer, df, sheetname):
    df.to_excel(writer, sheet_name=sheetname, index=False)
    worksheet = writer.sheets[sheetname]
    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((series.astype(str).map(len).max(), # len of largest item
                       len(str(series.name)) # len of header
                       )) + 1 # a bit of extra space
        worksheet.set_column(idx, idx, max_len)


class CoachList:
    def __init__(self, existing_fp=None):
        self.coaches = {}
        if existing_fp is not None:
            self.read_existing(existing_fp)

    def create_list(self, coachlist):
        for coach in coachlist:
            if coach not in self.coaches and len(coach) > 0:
                self.coaches[coach] = Coach(coach)
    
    def delete_all_but_listed(self, safelist):
        to_delete = []
        for coach in self.coaches.keys():
            if coach not in safelist:
                to_delete.append(coach)
        for coach in to_delete:
            del self.coaches[coach]

    def read_existing(self, fp):
        if os.path.exists(fp):
            self._read_existing(fp)

    def _read_existing(self, fp):
        xl_file = pd.ExcelFile(fp)
        coach_sheets = [a for a in xl_file.sheet_names if a not in ["Summary", "Mailing lists", "Currency", "Remits", "Providerships"]]
        for coach_name in coach_sheets:
            if coach_name not in self.coaches:
                self.coaches[coach_name] = Coach(coach_name)
            coach_info = xl_file.parse(coach_name)
            details = coach_info["Detail value"]
            details.index = coach_info["Detail type"]
            self.coaches[coach_name].name = details["Name"]
            self.coaches[coach_name].email_address = details["Email address"]
            self.coaches[coach_name].membership_status = details["BC membership status"]
            self.coaches[coach_name].membership_number = details["BC membership number"]
            if str(details["CPD expiry"]) not in ["nan", "NaT"]:
                self.coaches[coach_name].cpd_expiry = details["CPD expiry"].date()
            if str(details["First aid expiry"]) not in ["nan", "NaT"]:
                self.coaches[coach_name].first_aid_expiry = details["First aid expiry"].date()
            if str(details["Safeguarding expiry"]) not in ["nan", "NaT"]:
                self.coaches[coach_name].safeguarding_expiry = details["Safeguarding expiry"].date()
            if str(details["DBS date"]) not in ["nan", "NaT"]:
                self.coaches[coach_name].dbs_expiry = details["DBS date"].date()

            if "Remit craft" in coach_info:
                remit_details = coach_info["Remit craft"]
                remit_details.index = coach_info["Remit environment"]
                for a in remit_order:
                    self.coaches[coach_name].remits[a] = remit_details[a]

            if "Club sign off for" in coach_info:
                club_signoff_for = non_nan(coach_info["Club sign off for"])
                club_signoff_by = non_nan(coach_info["Club sign off by"])
                if len(club_signoff_for) != len(club_signoff_by):
                    raise ValueError("Signoff data for coach %s is corrupt"%coach_name)
                for i in range(len(club_signoff_for)):
                    self.coaches[coach_name].add_club_signoff(club_signoff_for[i], 
                                                              club_signoff_by[i])
            if "Qualification name" in coach_info:
                qual_names = non_nan(coach_info["Qualification name"])
                qual_types = non_nan(coach_info["Qualification type"])
                if len(qual_names) != len(qual_types):
                    raise ValueError("Qualification data for coach %s is corrupt"%coach_name)
                for i in range(len(qual_names)):
                    self.coaches[coach_name].add_qualification(qual_names[i],
                                                               qual_types[i])

            if "Safety training course" in coach_info:
                safety_training = non_nan(coach_info["Safety training course"])
                safety_date = non_nan_date(coach_info["Safety training date"])
                if len(safety_training) != len(safety_date):
                    raise ValueError("Safety training data for coach %s is corrupt"%coach_name)
                for i in range(len(safety_training)):
                    self.coaches[coach_name].add_safety_training(safety_training[i], 
                                                                 safety_date[i])
            
            if "First aid training type" in coach_info:
                first_aid_type = non_nan(coach_info["First aid training type"])
                first_aid_date = non_nan_date(coach_info["First aid training date"])
                if len(first_aid_type) != len(first_aid_date):
                    raise ValueError("First aid data for coach %s is corrupt"%coach_name)
                for i in range(len(first_aid_type)):
                    self.coaches[coach_name].add_first_aid_training(first_aid_type[i],
                                                                    first_aid_date[i])

            if "Safeguarding training type" in coach_info:
                safeguarding_types = non_nan(coach_info["Safeguarding training type"])
                safeguarding_expiry = non_nan_date(coach_info["Safeguarding training expiry"])
                safeguarding_dates = non_nan_date(coach_info["Safeguarding training date"])
                if len(safeguarding_types) != len(safeguarding_expiry) or\
                   len(safeguarding_types) != len(safeguarding_dates):
                    raise ValueError("Safeguarding data for coach %s is corrupt"%coach_name)
                for i in range(len(safeguarding_types)):
                    self.coaches[coach_name].add_safeguarding_training(safeguarding_types[i], 
                                                                       safeguarding_expiry[i],
                                                                       safeguarding_dates[i])

            if "Provider role" in coach_info:
                provider_types = non_nan(coach_info["Provider role"])
                provider_dates = non_nan_date(coach_info["Provider role date"])
                provider_active = non_nan(coach_info["Provider role active"])
                if len(provider_types) != len(provider_dates) or\
                   len(provider_dates) != len(provider_active):
                    raise ValueError("Provider data for coach %s is corrupt"%coach_name)
                for i in range(len(provider_types)):
                    self.coaches[coach_name].add_provider_credential(provider_types[i],
                                                                     provider_dates[i],
                                                                     provider_active[i])
            
    def read_qualifications(self, fp, safelist=None):
        if os.path.exists(fp):
            xl_file = pd.ExcelFile(fp)
            sheet_info = xl_file.parse("My Club Members with Coach Vali")
            entries = len(sheet_info)
            for i in range(entries):
                row = sheet_info.iloc[i]
                coach_name = row["Name"]
                if safelist:
                    if coach_name not in safelist:
                        continue
                if coach_name not in self.coaches:
                    self.coaches[coach_name] = Coach(coach_name)
                qual_name = row["Qualification Name"]
                qual_type = row["Qualification Category"]
                self.coaches[coach_name].add_qualification(qual_name, qual_type, True)
                self.coaches[coach_name].membership_status = row["Membership Status"]
                self.coaches[coach_name].safeguarding_expiry = row["Safeguarding From"].date()
                self.coaches[coach_name].first_aid_expiry = row["First Aid Expiry"].date()
                self.coaches[coach_name].cpd_expiry = row["CPD Expiry"].date()
    
    def read_credentials(self, fp, safelist=None):
        if os.path.exists(fp):
            xl_file = pd.ExcelFile(fp)
            sheet_info = xl_file.parse("Club All Credentials")
            entries = len(sheet_info)
            for i in range(entries):
                row = sheet_info.iloc[i]
                coach_name = row["Firstname"] + " " + row["Lastname"]
                if safelist:
                    if coach_name not in safelist:
                        continue
                if coach_name not in self.coaches:
                    self.coaches[coach_name] = Coach(coach_name)
                prov_cred = row["Training"]
                date = row["Completed On"].date()
                active = row["Status"]
                
                disqual_snippets = ["Training", "Orientation", "Paddlepower", "*"]
                disqual = True in [a in prov_cred for a in disqual_snippets]
                
                if (("Provider" in prov_cred) or ("PPA" in prov_cred)) and not disqual:
                    self.coaches[coach_name].add_provider_credential(prov_cred, date, active, True)

    def read_safety_report(self, fp, safelist=None):
        if os.path.exists(fp):
            xl_file = pd.ExcelFile(fp)
            sheet_info = xl_file.parse("Safety Report")
            entries = len(sheet_info)
            for i in range(entries):
                row = sheet_info.iloc[i]
                coach_name = row["Firstname"] + " " + row["Lastname"]
                if safelist:
                    if coach_name not in safelist:
                        continue
                if coach_name not in self.coaches:
                    self.coaches[coach_name] = Coach(coach_name)
                safety_training = row["Training"]
                date_completed = row["Completed On"].date()
                self.coaches[coach_name].add_safety_training(safety_training, date_completed, True)
                self.coaches[coach_name].membership_number = row["Membership Number"]

    def read_first_aid_report(self, fp, safelist=None):
        if os.path.exists(fp):
            xl_file = pd.ExcelFile(fp)
            sheet_info = xl_file.parse("First Aid Report")
            entries = len(sheet_info)
            for i in range(entries):
                row = sheet_info.iloc[i]
                coach_name = row["Firstname"] + " " + row["Lastname"]
                if safelist:
                    if coach_name not in safelist:
                        continue
                if coach_name not in self.coaches:
                    self.coaches[coach_name] = Coach(coach_name)
                name = row["Training"]
                date = row["Completed On"].date()
                self.coaches[coach_name].add_first_aid_training(name, date, True)
                self.coaches[coach_name].membership_number = row["Membership Number"]

    def read_safeguarding_report(self, fp, safelist=None):
        if os.path.exists(fp):
            xl_file = pd.ExcelFile(fp)
            sheet_info = xl_file.parse("Safeguarding Clubs Report")
            entries = len(sheet_info)
            for i in range(entries):
                row = sheet_info.iloc[i]
                coach_name = row["Firstname"] + " " + row["Lastname"]
                if safelist:
                    if coach_name not in safelist:
                        continue
                if coach_name not in self.coaches:
                    self.coaches[coach_name] = Coach(coach_name)
                name = row["Name"]
                expiry = row["Expiry Date"].date()
                date = row["Granted Date"].date()
                self.coaches[coach_name].add_safeguarding_training(name, expiry, date, True)
                self.coaches[coach_name].membership_number = row["rpt Members MID"]

    def produce_currency_dataframe(self):
        data = []
        headings = ["Name", "BC membership", "CPD expiry", 
                    "First aid expiry", "Safeguarding expiry", "DBS date"]
        for coach in sorted(self.coaches.keys()):
            data.append([self.coaches[coach].name,
                         self.coaches[coach].membership_status,
                         self.coaches[coach].cpd_expiry,
                         self.coaches[coach].first_aid_expiry,
                         self.coaches[coach].safeguarding_expiry,
                         self.coaches[coach].dbs_expiry])
        return pd.DataFrame(data, columns=headings)

    def produce_remit_dataframe(self):
        data = []
        headings = ["Coach name"] + remit_order
        for coach in sorted(self.coaches.keys()):
            coach_data = [coach] + [self.coaches[coach].remits[a] for a in remit_order]
            data.append(coach_data)
        return pd.DataFrame(data, columns=headings)
    
    def produce_provider_dataframe(self):
        data = []
        headings = ["Coach name"] + provider_options
        for coach in sorted(self.coaches.keys()):
            # create blank row, then fill
            coach_data = [coach] + ["" for a in provider_options]
            
            # check each of their credentials in turn annd fill in as appropriate
            for cred in self.coaches[coach].provider_credentials:
                # does this make them a valid ppa provider 
                # rules for this are different to other statuses
                if cred.name in ["PPA Provider eLearning", "PPA Moderation eLearning"]:
                    coach_data[1] = cred.active
                
                # check credential against
                for i, s in enumerate(provider_options[1:]):
                    if cred.name == s:
                        coach_data[2+i] = cred.active
            
            data.append(coach_data)
        return pd.DataFrame(data, columns=headings)
                
    def write_to_excel(self, fp):
        with pd.ExcelWriter(fp) as writer:
            cr_df = self.produce_currency_dataframe()
            add_to_writer_with_fixed_col_widths(writer, cr_df, sheetname="Currency")

            rm_df = self.produce_remit_dataframe()
            add_to_writer_with_fixed_col_widths(writer, rm_df, sheetname="Remits")
            
            pv_df = self.produce_provider_dataframe()
            add_to_writer_with_fixed_col_widths(writer, pv_df, sheetname="Providerships")
            
            for coach in sorted(self.coaches.keys()):
                coach_df = self.coaches[coach].produce_dataframe()
                add_to_writer_with_fixed_col_widths(writer, coach_df, sheetname=coach)

            
class Coach:
    def __init__(self, name):
        self.name = name
        self.email_address = None
        self.membership_status = None
        self.membership_number = None
        self.cpd_expiry = None
        self.first_aid_expiry = None
        self.safeguarding_expiry = None
        self.dbs_expiry = None
        self.remits = {a: "" for a in remit_order}
        self.club_signoffs = []
        self.qualifications = []
        self.safety_training = []
        self.first_aid_training = []
        self.safeguarding_training = []
        self.provider_credentials = []

    def add_club_signoff(self, sign_off_for, sign_off_by):
        new_report = ClubSignOff(sign_off_for, sign_off_by)
        if new_report not in self.club_signoffs:
            self.club_signoffs.append(new_report)

    def add_qualification(self, qualification, qual_type, verbose=False):
        if qual_type in ["Coaching Award", "Leadership Award", "Performance Award"]:
            new_report = Qualification(qualification, qual_type)
            if new_report not in self.qualifications:
                self.qualifications.append(new_report)
                if verbose:
                    print("New qualification for %s:\n%s\n"%(self.name, new_report))

    def add_safety_training(self, training, date, verbose=False):
        new_report = SafetyTraining(training, date)
        if new_report not in self.safety_training:
            self.safety_training.append(new_report)
            if verbose:
                print("New safety training for %s\n%s\n"%(self.name, new_report))

    def add_first_aid_training(self, training, date, verbose=False):
        new_report = FirstAidTraining(training, date)
        if new_report not in self.first_aid_training:
            self.first_aid_training.append(new_report)
            if verbose:
                print("New first aid for %s\n%s\n"%(self.name, new_report))

    def add_safeguarding_training(self, name, expiry, date, verbose=False):
        new_report = SafeguardingTraining(name, expiry, date)
        if new_report not in self.safeguarding_training:
            self.safeguarding_training.append(new_report)
            if verbose:
                print("New safeguarding for %s\n%s\n"%(self.name, new_report))
                
    def add_provider_credential(self, name, date, active, verbose=False):
        new_report = ProviderCredential(name, date, active)
        if new_report not in self.provider_credentials:
            self.provider_credentials.append(new_report)
            if verbose:
                print("New provider credential for %s\n%s\n"%(self.name, new_report))
                    
    def produce_dataframe(self):
        frames = []

        detail_data = [["Name", self.name], 
                       ["Email address", self.email_address],
                       ["BC membership number", self.membership_number],
                       ["BC membership status", self.membership_status],
                       ["CPD expiry", self.cpd_expiry],
                       ["First aid expiry", self.first_aid_expiry],
                       ["Safeguarding expiry", self.safeguarding_expiry],
                       ["DBS date", self.dbs_expiry]]
        detail_headings = ["Detail type", "Detail value"]
        frames.append(pd.DataFrame(detail_data, columns=detail_headings))

        remit_data = [[a, self.remits[a]] for a in remit_order]
        frames.append(pd.DataFrame(remit_data, columns=["Remit environment", "Remit craft"]))

        signoff_data = [[a.sign_off_for, a.sign_off_by] for a in self.club_signoffs]
        signoff_headings = ["Club sign off for",  "Club sign off by"]
        frames.append(pd.DataFrame(signoff_data, columns=signoff_headings))

        qualification_data = [[a.name, a.type] for a in self.qualifications]
        qualification_headings = ["Qualification name", "Qualification type"]
        frames.append(pd.DataFrame(qualification_data, columns=qualification_headings))

        safety_data = [[a.name, a.date] for a in self.safety_training]
        safety_headings = ["Safety training course", "Safety training date"]
        frames.append(pd.DataFrame(safety_data, columns=safety_headings))

        first_aid_data = [[a.name, a.date] for a in self.first_aid_training]
        first_aid_headings = ["First aid training type", "First aid training date"]
        frames.append(pd.DataFrame(first_aid_data, columns=first_aid_headings))

        safeguarding_data = [[a.name, a.expiry, a.date] for a in self.safeguarding_training]
        safeguarding_headings = ["Safeguarding training type", 
                                 "Safeguarding training expiry", 
                                 "Safeguarding training date"]
        frames.append(pd.DataFrame(safeguarding_data, columns=safeguarding_headings))
        
        provider_data = [[a.name, a.date, a.active] for a in self.provider_credentials]
        provider_headings = ["Provider role", "Provider role date", "Provider role active"]
        frames.append(pd.DataFrame(provider_data, columns=provider_headings))

        return pd.concat(frames, axis=1)

    def __str__(self):
        ret_str = "Name: %s\n" % self.name
        ret_str += "Email address: %s\n" % self.email_address
        ret_str += "BC Membership number: %s\n" % self.membership_number
        ret_str += "BC Membership status: %s\n" % self.membership_status
        ret_str += "CPD Expiry: %s\n" % self.cpd_expiry
        ret_str += "First Aid Expiry: %s\n" % self.first_aid_expiry
        ret_str += "Safeguarding Expiry: %s\n" % self.safeguarding_expiry
        ret_str += "DBS Expiry: %s\n" % self.dbs_expiry
        
        ret_str += "\nRemits: %i\n"%sum([1 for a in self.remits if self.remits[a]])
        for a in remit_order:
            if self.remits[a]:
                ret_str += "%s %s\n" % (a, self.remits[a])
        
        ret_str += "\nClub Signoffs: %i\n"%len(self.club_signoffs)
        for a in self.club_signoffs:
            ret_str += "%s\n" % a
        
        ret_str += "\nQualifications: %i\n"%len(self.qualifications)
        for a in self.qualifications:
            ret_str += "%s\n" % a
        
        ret_str += "\nSafety Training: %i\n"%len(self.safety_training)
        for a in self.safety_training:
            ret_str += "%s\n" % a
        
        ret_str += "\nFirst Aid Training: %i\n"%len(self.first_aid_training)
        for a in self.first_aid_training:
            ret_str += "%s\n" % a
        
        ret_str += "\nSafeguarding Training: %i\n"%len(self.safeguarding_training)
        for a in self.safeguarding_training:
            ret_str += "%s\n" % a
        
        ret_str += "\nProvider credentials: %i\n"%len(self.provider_credentials)
        for a in self.provider_credentials:
            ret_str += "%s\n" % a

        return ret_str


class ProviderCredential:
    def __init__(self, name, date, active):
        self.name = name
        self.date = date
        self.active = active

    def __eq__(self, a):
        name_same = self.name == a.name
        date_same = self.date == a.date
        active_same = self.active == a.active
        return name_same and date_same and active_same
    
    def __str__(self):
            return "%s %s %s"%(self.name, self.date, self.active)


class ClubSignOff:
    def __init__(self, sign_off_for, sign_off_by):
        self.sign_off_for = sign_off_for
        self.sign_off_by = sign_off_by

    def __eq__(self, a):
        sof_same = self.sign_off_for == a.sign_off_for
        sob_same = self.sign_off_by == a.sign_off_by
        return sof_same and sob_same

    def __str__(self):
        return "%s %s"%(self.sign_off_for, self.sign_off_by)


class Qualification:
    def __init__(self, name, qual_type, date=None):
        self.name = name
        self.type = qual_type
        self.date = date

    def __eq__(self, a):
        name_same = self.name == a.name
        type_same = self.type == a.type
        date_same = self.date == a.date
        return name_same and type_same and date_same
    
    def __str__(self):
        if self.date is None:
            return "%s %s"%(self.name, self.type)
        else:
            return "%s %s %s"%(self.name, self.type, self.date)


class SafetyTraining:
    def __init__(self, name, date):
        self.name = name
        self.date = date

    def __eq__(self, a):
        return (self.name == a.name) and (self.date == a.date)

    def __str__(self):
        return "%s %s"%(self.name, self.date)


class FirstAidTraining:
    def __init__(self, name, date):
        self.name = name
        self.date = date

    def __eq__(self, a):
        return (self.name == a.name) and (self.date == a.date)

    def __str__(self):
        return "%s %s"%(self.name, self.date)


class SafeguardingTraining:
    def __init__(self, name, expiry_date, date):
        self.name = name
        self.date = date
        if str(expiry_date) == "NaT":
            self.expiry = "-"
        else:
            self.expiry = expiry_date

    def __eq__(self, a):
        name_same = self.name == a.name
        expiry_same = self.expiry == a.expiry
        date_same = self.date == a.date
        return name_same and expiry_same and date_same

    def __str__(self):
        return "%s %s %s"%(self.name, self.expiry, self.date)


if __name__ == "__main__":
    coach_names = []
    with open("active_coaches.txt") as inFile:
        for line in inFile:
            coach_names.append(" ".join(line.split()))

    existing_fp = sys.argv[1] if len(sys.argv) > 1 else "Coaches Summary.xlsx"
    cl = CoachList()
    cl.read_existing(existing_fp)
    cl.read_safety_report("Safety Report.xlsx", coach_names)
    cl.read_first_aid_report("First Aid Training.xlsx", coach_names)
    cl.read_safeguarding_report("Safeguarding Report.xlsx", coach_names)
    cl.read_qualifications("My Club Members with Coach Validation.xlsx", coach_names)
    cl.read_credentials("All Member Credentials.xlsx", coach_names)

    #print("Total entries: %i"%len(cl.coaches))
    #print(", ".join(cl.coaches.keys()))

    cl.delete_all_but_listed(coach_names)
    cl.create_list(coach_names)
    
    #print(cl.coaches.keys())
    cl.write_to_excel("Coaches Summary.xlsx")

    #print("\nExample details of one coach:")
    #print(cl.coaches["Alex Allen"])

    #print("---")

