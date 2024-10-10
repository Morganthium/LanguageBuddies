from fuzzywuzzy import process
import pandas as pd

class Participant:
    def __init__(self, name, email, fluent_languages, learning_languages):
        self.name = name
        self.email = email
        self.fluent_languages = fluent_languages
        self.learning_languages = learning_languages
        self.matched_with = []

    def can_match(self, other):
        if other.name in self.matched_with or len(self.matched_with) >= 3:
            return False, None, None
        for ll1 in self.learning_languages:
            if ll1 in other.learning_languages:
                return True, "Learning", ll1
        return False, None, None

    def match(self, other, match_type, language):
        self.matched_with.append((other.name, match_type, language))
        other.matched_with.append((self.name, match_type, language))


def standardize_language(input_language, standard_languages, threshold=90):
    # Use fuzzy matching to find the closest standard language name
    match, score = process.extractOne(input_language, standard_languages)
    if score > threshold:
        return match
    else:
        # If no close match is found, return the input for manual review
        return "Review: " + input_language

# Define a list of standardized language names
standard_languages = ["English", "Spanish", "French", "German", "Chinese", "Japanese", "Korean", "Italian", "Cantonese", "Russian"]  # Complete this list based on your needs

def load_and_organize_data(file_name):
    df = pd.read_excel(file_name, engine="openpyxl")
    participants_by_language = {}
    for _, row in df.iterrows():
        learning_languages = [standardize_language(row['Language you want to learn 1'], standard_languages)]
        if pd.notna(row['Language you want to learn 2 (Optional)']):
            learning_languages.append(standardize_language(row['Language you want to learn 2 (Optional)'], standard_languages))

        participant = Participant(row['Name'], row['Student Email'], [], learning_languages)
        
        for language in learning_languages:
            if language not in participants_by_language:
                participants_by_language[language] = []
            participants_by_language[language].append(participant)
    
    return participants_by_language


participants_by_language = load_and_organize_data("Language_Buddies_Sign_Up_(Responses).xlsx")


def find_matches(participants_by_language):
    unmatched = []
    matched_pairs = set()  # To keep track of matched pairs and avoid duplicates

    for language, learners in participants_by_language.items():
        matched_this_round = set()  # Keep track of participants matched in this iteration
        for i, p1 in enumerate(learners):
            for j, p2 in enumerate(learners[i+1:], start=i+1):
                if p1.name == p2.name or (p1.name, p2.name) in matched_pairs or (p2.name, p1.name) in matched_pairs:
                    continue  # Skip if they are the same or already matched
                can_match, match_type, matched_language = p1.can_match(p2)
                if can_match:
                    p1.match(p2, match_type, matched_language)
                    matched_pairs.add((p1.name, p2.name))  # Add to matched pairs to avoid duplicates
                    matched_this_round.add(p1)
                    matched_this_round.add(p2)
                    break  # Move to the next participant after a successful match
        
        # Update learners list by removing matched participants
        learners[:] = [p for p in learners if p not in matched_this_round]
        unmatched.extend(learners)  # Add remaining unmatched participants

    return unmatched



# Update the SAVE_MATCHES function
def save_matches_to_file(participants_by_language, unmatched, output_filename):
    rows = []
    for language, participants in participants_by_language.items():
        for participant in participants:
            for match_name, match_type, match_language in participant.matched_with:
                matched_participant = next((p for p in participants if p.name == match_name), None)
                if matched_participant:
                    rows.append([participant.name, participant.email, matched_participant.name, matched_participant.email, "Both are learning", match_language])

    # Add unmatched participants
    for participant in unmatched:
        rows.append([participant.name, participant.email, "Unmatched", "", "Learning", ""])

    df_output = pd.DataFrame(rows, columns=["Name", "Email", "Matched Name", "Matched Email", "Status", "Language"])
    df_output.to_excel(output_filename, index=False)

# Adjusted function call
participants_by_language = load_and_organize_data("Language_Buddies_Sign_Up_(Responses).xlsx")
unmatched = find_matches(participants_by_language)
save_matches_to_file(participants_by_language, unmatched, "matched_participants.xlsx")
