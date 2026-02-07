import * as React from 'react';
import { useState } from 'react';
import { Button } from '@fluentui/react-button';

export interface IActionCardProps {
    step: string;
    action: string;
    userName: string;
    userImage: string;
    comments: string;
    actionDate: string;
    actioncolorCode: string;
    arrowdirection: string;
    surveyResponseId: string | number | null;  // Refined the type for surveyResponseId
}

export const ActionCard: React.FC<IActionCardProps> = ({
    step,
    action,
    userName,
    userImage,
    comments,
    actionDate,
    actioncolorCode,
    surveyResponseId,
    arrowdirection
}) => {
    const [isHovered, setIsHovered] = useState(false);  // State to track hover effect

    const handleMouseEnter = () => setIsHovered(true);
    const handleMouseLeave = () => setIsHovered(false);
    const getDirection = () => {
        if (Xrm.Utility.getGlobalContext().userSettings.languageId === 1025)
            return 'rtl';
        return 'ltr'
    };
    // This function no longer returns a Promise directly, ensuring no async directly in the onClick
    const openSurveyResponseModal = (surveyResponseId: string | number | null) => {
        if (!surveyResponseId) return;  // Early return if surveyResponseId is null or undefined

        const Data = {
            ResponseId: surveyResponseId
        };

        const pageInput: Xrm.Navigation.PageInputHtmlWebResource = {
            pageType: 'webresource',
            webresourceName: "duc_Survey_View_HTML",  // Make sure this is properly defined elsewhere
            data: JSON.stringify(Data)
        };

        const navigationOptions: Xrm.Navigation.NavigationOptions = {
            target: 2, // Modal dialog
            width: { value: 700, unit: 'px' },
            height: { value: 1000, unit: 'px' },
            position: 1 // Center
        };

        // Use .then and .catch here to handle the promise
        Xrm.Navigation.navigateTo(pageInput, navigationOptions)
            .then(() => {
                console.log("Survey response opened successfully.");
                return undefined;  // Explicitly return undefined to satisfy the "always-return" rule
            })
            .catch((error: any) => {
                console.error("Error opening survey response view:", error);
                return undefined;  // Explicitly return undefined to satisfy the "always-return" rule
            });
    };

    return (
        <div
            className="action-card"
            onMouseEnter={handleMouseEnter}
            onMouseLeave={handleMouseLeave}
            style={{
                boxShadow: `0 2px 6px ${actioncolorCode}`,
                direction: getDirection()
            }}
        >
            {/* Card header with Step and Action */}
            <div className="action-card-header">
                <h3>{step}</h3>
            </div>

            {/* Card body with image */}
            <div className="action-card-body">
                <img
                    src={userImage}
                    className="user-image"
                />
                {/* On hover, display comments and survey button */}
                {isHovered && (
                    <div className="card-overlay visible">
                        <div className="card-comments">
                            <p>{comments}</p>
                        </div>
                        <div>
                            {/* Show survey button only if surveyResponseId exists */}
                            {surveyResponseId && (
                                <Button
                                    onClick={() => openSurveyResponseModal(surveyResponseId)} // Avoid async/await here
                                >
                                    View Survey
                                </Button>
                            )}
                        </div>
                    </div>
                )}
            </div>


            {/* Card footer with Action Date */}
            <div className="action-card-footer" style={{
                backgroundColor: actioncolorCode
            }}>
                <p>{userName}</p>
                <p>{actionDate}</p>
                <p>{action}</p>
            </div>
            <div className={`arrow-line ${arrowdirection}`}></div>
        </div>
    );
};
