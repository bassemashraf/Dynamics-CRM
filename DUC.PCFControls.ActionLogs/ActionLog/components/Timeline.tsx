import * as React from 'react';
import { useState, useEffect } from 'react';
import { ActionCard } from './ActionCard';

export interface ITimelineProps {
    records: Array<any>;  // Array of records
}

export const Timeline: React.FC<ITimelineProps> = ({ records }) => {
    const [numberOfCardsPerRow, setNumberOfCardsPerRow] = useState(4);  // Default to 4 cards per row

    useEffect(() => {
        const handleResize = () => {
            // Adjust number of cards per row based on window width
            if (window.innerWidth < 600) {
                setNumberOfCardsPerRow(1);  // 1 card per row for small screens
            } else if (window.innerWidth < 900) {
                setNumberOfCardsPerRow(2);  // 2 cards per row for medium screens
            } else if (window.innerWidth < 1200) {
                setNumberOfCardsPerRow(3);  // 3 cards per row for large screens
            } else {
                setNumberOfCardsPerRow(4);  // 4 cards per row for extra large screens
            }
        };

        // Add resize listener
        window.addEventListener('resize', handleResize);

        // Initial call to set the number of cards
        handleResize();

        // Clean up event listener on component unmount
        return () => window.removeEventListener('resize', handleResize);
    }, []);

    const rows = [];
    // Split records into rows based on the number of cards per row
    for (let i = 0; i < records.length; i += numberOfCardsPerRow) {
        rows.push(records.slice(i, i + numberOfCardsPerRow));
    }
    const getDirection = () => {
        if (Xrm.Utility.getGlobalContext().userSettings.languageId === 1025)
            return 'rtl';
        return 'ltr'
    };
    const getOppositeDirection = () => {
        if (getDirection() == 'ltr')
            return 'rtl';
        return 'ltr'
    }
    const IsEnglish = () => {
        if (Xrm.Utility.getGlobalContext().userSettings.languageId === 1025)
            return false;
        return true;
    }
    return (
        <div className="timeline-container">
            {rows.map((row, rowIndex) => {
                const isEvenRow = rowIndex % 2 === 1;  // Even row (index starts from 0)
                return (
                    <div
                        key={rowIndex}
                        className={`timeline-row ${isEvenRow ? getOppositeDirection() : getDirection()}`}
                    >
                        {row.map((record, index) => {
                            const StepNumber = (index + 1) + numberOfCardsPerRow * rowIndex;
                            const step = StepNumber.toString() + "."
                                + (IsEnglish()
                                    ? record.getFormattedValue("duc_processstage")
                                    : record.getFormattedValue("a_be9b6fc5daff400b925124addbc7ebff.duc_arabicname"));

                            const action = IsEnglish()
                                ? record.getFormattedValue("duc_action")
                                : record.getFormattedValue("a_ed0ef68caced43b082fbd1bb9c8ff0f8.duc_arabicname");

                            const userName = IsEnglish()
                                ? record.getFormattedValue("ownerid")
                                : record.getFormattedValue("a_dd94068753c14a77a4837afc09c11147.duc_usernamearabic");

                            const comments = record.getFormattedValue("duc_comments");
                            const actionDate = record.getFormattedValue("createdon");
                            const userImage = record.getFormattedValue("a_dd94068753c14a77a4837afc09c11147.photourl");
                            const userEntityImage = "";//record.getFormattedValue("a_dd94068753c14a77a4837afc09c11147.domainname");
                            /*const surveyResponseId = record.getFormattedValue("duc_surveyresponse") ? record._record.fields["duc_surveyresponse"].reference.id.guid : null;//record.getFormattedValue("duc_surveyresponse");*/
                            const actioncolorCode = record.getFormattedValue("a_4b963fcb00d44804b6b8b0c07ce8917f.duc_color");
                            let arrowDirection = getDirection();

                            if ((index == row.length - 1 && rowIndex == rows.length - 1) || (index < row.length - 1))
                                arrowDirection = isEvenRow ? getOppositeDirection() : getDirection();
                            else if (index == row.length - 1 && rowIndex < rows.length - 1)
                                arrowDirection = 'down';



                            return (
                                <div key={record.getRecordId()} className="card-wrapper">
                                    <ActionCard
                                        step={step || 'N/A'}
                                        action={action || 'No Action'}
                                        userImage={'/_imgs/contactphoto.png'}
                                        userName={userName || 'Unknown User'}
                                        comments={comments || ''}
                                        actionDate={actionDate || 'No Date'}
                                        actioncolorCode={actioncolorCode || 'rgba(0, 0, 0, 0.1)'}
                                        surveyResponseId={null}
                                        arrowdirection={arrowDirection || 'ltr'}
                                    />

                                </div>
                            );
                        })}
                    </div>
                );
            })}
        </div>
    );
};

