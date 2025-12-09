"""
Neo4j Data Extractor for Museum Objects
Extracts structured object information from Neo4j database to Excel
"""

from story_generator import Neo4jExtractor

def main():
    """Extract museum object data from Neo4j to Excel"""
    
    # Configure   Neo4j  
    extractor = Neo4jExtractor(
        url="bolt://localhost:7687",
        username="neo4j",
        password="neo4j"  # put your password
    )
    
    # Extract 100 objects to Excel
    output_file = extractor.extract_to_excel(
        output_file="catalog_100_objects.xlsx",
        limit=100
    )
    
    print(f"\n✅ Data extraction complete!")
    print(f"📁 File: {output_file}")
    print(f"\n💡 Next step: Generate stories using story_generator.py")

if __name__ == "__main__":
    main()
