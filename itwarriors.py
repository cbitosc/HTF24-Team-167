import pandas as pd
import openpyxl

class PublicationSummaryGenerator:
    def __init__(self, file_path, file_type):
       
        self.file_path = file_path
        self.file_type = file_type
        self.publications_df = self.load_publications()

    def load_publications(self):
        
        if self.file_type == 'excel':
            publications_df = pd.read_excel(self.file_path)
        elif self.file_type == 'bibtex':
            print("BibTeX loading not implemented.")
            return pd.DataFrame()
        else:
            raise ValueError("Unsupported file type. Use 'excel' or 'bibtex'.")
        
        return publications_df



    def filter_by_title(self, title_keyword):
        filtered_df = self.publications_df[self.publications_df['title'].str.contains(title_keyword, case=False, na=False)]
        self.export_to_excel("Filtered_By_Title.xlsx", filtered_df, sheet_name="Filtered By Title")
        return filtered_df

    def filter_by_author(self, author_name):
        filtered_df = self.publications_df[self.publications_df['author'].str.contains(author_name, case=False, na=False)]
        self.export_to_excel("Filtered_By_Author.xlsx", filtered_df, sheet_name="Filtered By Author")
        return filtered_df

    def filter_by_year(self, start_year, end_year):
        filtered_df = self.publications_df[
            (self.publications_df['year'] >= start_year) & (self.publications_df['year'] <= end_year)
        ]
        self.export_to_excel("Filtered_By_Year_Range.xlsx", filtered_df, sheet_name="Filtered By Year Range")
        return filtered_df

    def filter_by_type(self, pub_type):
        filtered_df = self.publications_df[self.publications_df['type'].str.contains(pub_type, case=False, na=False)]
        self.export_to_excel("Filtered_By_Type.xlsx", filtered_df, sheet_name="Filtered By Type")
        return filtered_df

    def generate_yearly_summary(self):
        summary_df = self.publications_df.groupby(['year', 'type']).size().unstack(fill_value=0)
        return summary_df

    def generate_total_summary(self):
        total_summary_df = self.publications_df.groupby('year').size().reset_index(name='total_publications')
        return total_summary_df


    def export_to_excel(self, output_path, summary_df, sheet_name="Publication Summary"):
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name=sheet_name)
        print(f"Excel file saved to {output_path}")


def display_menu():
    print("\nPublication Summary Generator - Filter Menu")
    print("1. Filter by Title")
    print("2. Filter by Author")
    print("3. Filter by Year Range")
    print("4. Filter by Type")
    print("5. Generate Yearly Summary")
    print("6. Generate Total Summary")
    print("7. Exit")
    return input("Enter your choice (1-7): ")



if __name__ == "__main__":
    generator = PublicationSummaryGenerator(file_path='publications.xlsx', file_type='excel')
    while True:
        choice = display_menu()
        if choice == '1':
            title = input("Enter the title keyword to filter by: ")
            generator.filter_by_title(title)

        elif choice == '2':
            author_name = input("Enter the author name to filter by: ")
            generator.filter_by_author(author_name)

        elif choice == '3':
            start_year = int(input("Enter start year: "))
            end_year = int(input("Enter end year: "))
            generator.filter_by_year(start_year, end_year)

        elif choice == '4':
            pub_type = input("Enter the type of publication (journal or conference): ")
            generator.filter_by_type(pub_type)

        elif choice == '5':
            yearly_summary = generator.generate_yearly_summary()
            print("Yearly Summary:\n", yearly_summary)

        elif choice == '6':
            total_summary = generator.generate_total_summary()
            print("Total Summary:\n", total_summary)

        elif choice == '7':
            print("Exiting the program.")
            break

        else:
            print("Invalid choice. Please enter a number between 1 and 7.")