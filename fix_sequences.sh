#!/bin/bash

# Fix PostgreSQL sequences after SQLite import
# This script resets all auto-increment sequences to the correct values

echo "Fixing PostgreSQL sequences..."

# Get the PostgreSQL container ID
CONTAINER_ID=$(docker ps -qf "name=hub_postgres")

if [ -z "$CONTAINER_ID" ]; then
    echo "Error: PostgreSQL container not found"
    exit 1
fi

echo "Found PostgreSQL container: $CONTAINER_ID"

# Fix sequences for all tables
docker exec -i $CONTAINER_ID psql -U hub_user -d hub_db <<EOF
-- Fix users table sequence
SELECT setval('users_id_seq', (SELECT MAX(id) FROM users));

-- Fix files table sequence
SELECT setval('files_id_seq', (SELECT MAX(id) FROM files));

-- Fix validation_results table sequence
SELECT setval('validation_results_id_seq', (SELECT MAX(id) FROM validation_results));

-- Fix macro_processing table sequence
SELECT setval('macro_processing_id_seq', (SELECT MAX(id) FROM macro_processing));

-- Verify sequences
SELECT 'users_id_seq: ' || last_value FROM users_id_seq;
SELECT 'files_id_seq: ' || last_value FROM files_id_seq;
SELECT 'validation_results_id_seq: ' || last_value FROM validation_results_id_seq;
SELECT 'macro_processing_id_seq: ' || last_value FROM macro_processing_id_seq;
EOF

echo "Done! Sequences have been reset."
